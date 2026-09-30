"""
powerfab.py
MIW client for the Tekla PowerFab Open API.

The API itself is .NET only, so this module does not talk to PowerFab directly. It builds
the XML request, shells out to MIWFab.exe (the one small C# bridge), and parses the XML
response back into rows you can put straight into a CSV or a DataFrame.

Everything above the bridge is Python. The bridge never changes.

    from powerfab import PowerFab, cmd

    pf = PowerFab()                      # reads powerfab.ini
    resp = pf.run(
        cmd("ProductionShippingProductivity_Get", GetOptions={
            "StatusType": "productivity-production-station",
            "DateSummary": "day",
            "StartDate":   "2026-08-14",
            "EndDate":     "2026-08-28",
            "GroupBy":     ["Job #", "Station"],
        })
    )
    rows = resp.rows("ProductionShippingProductivity_Get")

Design rules, every one of them learned by getting it wrong first:

  * Requests are validated against Tekla's SHIPPED XSD before they leave the machine.
    The developer website is a guide, not a spec, and in one morning it was wrong about
    three separate things: the class name is not the element name (fsreqProductionShipping
    ProductivityGet goes on the wire as ProductionShippingProductivity_Get), ConnectionGUID
    is an element and not an attribute, and the enum members it publishes are not the
    strings the schema accepts (it says productivityproductionstation, the schema says
    productivity-production-station). Run build_schema.py against the XSD in the PowerFab
    install to produce powerfab_schema.json, and none of that can bite again.

  * Children are REORDERED into schema sequence order automatically. The XSD uses
    xsd:sequence, so order is part of the contract, and keyword arguments in Python are
    not going to arrive in the right one by luck.

  * An XMLWarning is treated as a FAILURE. PowerFab silently drops an element it does not
    recognise and runs the command anyway, so a typo in a filter name returns an unfiltered
    dataset that looks perfectly plausible on a board. Never relax this.

  * Every command gets a fresh CommandGUID. On writes it is the only thing stopping a
    retry from double posting a cut or a receipt.
"""

from __future__ import annotations

import configparser
import csv
import json
import os
import re
import subprocess
import tempfile
import uuid
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Any

_HERE = os.path.dirname(os.path.abspath(__file__))


class PowerFabError(RuntimeError):
    pass


# --------------------------------------------------------------------------- schema


def _load_schema() -> dict:
    """
    powerfab_schema.json, generated from Tekla's shipped fabsuite-xml-request-v0108.xsd
    by build_schema.py. Absent, the client still works but validates nothing.
    """
    try:
        with open(os.path.join(_HERE, "powerfab_schema.json"), encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {"commands": {}, "types": {}}


SCHEMA = _load_schema()
COMMANDS: dict = SCHEMA.get("commands", {})
TYPES: dict = SCHEMA.get("types", {})


def _norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _suggest(bad: str, options) -> str:
    n = _norm(bad)
    exact = [o for o in options if _norm(o) == n]
    if exact:
        return f" Did you mean {exact[0]!r}?"
    near = [o for o in options if n and (n in _norm(o) or _norm(o) in n)]
    if near:
        return f" Close matches: {sorted(near)[:5]}"
    return f" Valid values: {sorted(options)[:12]}"


def describe(command: str) -> str:
    """Print what a command accepts, straight from the schema. Use it instead of guessing."""
    spec = COMMANDS.get(command)
    if spec is None:
        raise PowerFabError(f"{command!r} is not a PowerFab command.{_suggest(command, COMMANDS)}")
    out = [f"{command}   ({spec.get('type')})"]
    if spec.get("doc"):
        out.append(f"  {spec['doc']}")

    def walk(fields, indent="  "):
        for f in fields:
            flag = "required" if f["required"] else "optional"
            rep = ", repeatable" if f["repeatable"] else ""
            out.append(f"{indent}<{f['name']}>  {flag}{rep}")
            if f.get("doc"):
                out.append(f"{indent}    {f['doc']}")
            if f.get("enum"):
                for v in f["enum"]:
                    out.append(f"{indent}    = {v}")
            sub = TYPES.get(f.get("type"))
            if sub:
                walk(sub, indent + "    ")

    walk(spec.get("fields", []))
    return "\n".join(out)


# --------------------------------------------------------------------- request build


def _validate_and_order(el: ET.Element, fields: list[dict] | None, path: str) -> None:
    """Check names, enums and required children, then sort into schema sequence order."""
    if not fields:
        return
    by_name = {f["name"]: f for f in fields}

    for child in list(el):
        spec = by_name.get(child.tag)
        if spec is None:
            raise PowerFabError(
                f"<{child.tag}> is not valid inside {path}."
                f"{_suggest(child.tag, by_name.keys())}"
            )
        if spec.get("enum") and len(child) == 0:
            value = (child.text or "").strip()
            if value not in spec["enum"]:
                raise PowerFabError(
                    f"{path}/{child.tag}: {value!r} is not a legal value."
                    f"{_suggest(value, spec['enum'])}"
                )
        _validate_and_order(child, TYPES.get(spec.get("type")), f"{path}/{child.tag}")

    for spec in fields:
        if spec["required"] and el.find(spec["name"]) is None:
            raise PowerFabError(f"{path} is missing required element <{spec['name']}>.")

    order = {f["name"]: i for i, f in enumerate(fields)}
    kids = sorted(list(el), key=lambda c: order.get(c.tag, len(order)))
    for c in list(el):
        el.remove(c)
    for c in kids:
        el.append(c)


def cmd(name: str, _repeat: dict[str, list[Any]] | None = None, **fields: Any) -> ET.Element:
    """
    Build one command element, validated and ordered against the shipped schema.

        cmd("GetBillOfMaterials", JobNumber="2381",
            _repeat={"RequestedField": ["Weight", "Quantity"]})

    Nested dicts become nested elements, which is how every GetOptions block works.
    Lists repeat an element. A fresh CommandGUID is added automatically.
    """
    if COMMANDS and name not in COMMANDS:
        raise PowerFabError(
            f"{name!r} is not a PowerFab command element name.{_suggest(name, COMMANDS)}"
        )

    el = ET.Element(name)
    ET.SubElement(el, "CommandGUID").text = str(uuid.uuid4())

    def add(parent: ET.Element, key: str, value: Any) -> None:
        if isinstance(value, dict):
            child = ET.SubElement(parent, key)
            for k, v in value.items():
                add(child, k, v)
        elif isinstance(value, (list, tuple)):
            for v in value:
                add(parent, key, v)
        elif isinstance(value, bool):
            ET.SubElement(parent, key).text = "1" if value else "0"
        elif value is not None:
            ET.SubElement(parent, key).text = str(value)

    for k, v in fields.items():
        add(el, k, v)
    for k, vs in (_repeat or {}).items():
        for v in vs:
            add(el, k, v)

    _validate_and_order(el, COMMANDS.get(name, {}).get("fields"), name)
    return el


def build_request(*commands: ET.Element) -> str:
    """
    Wrap commands in a FabSuiteXMLRequest root.

    Every command on the root is a collection, so batching many commands into one document
    is both legal and the right thing to do. They execute in order and return in order, in
    a single round trip.
    """
    root = ET.Element("FabSuiteXMLRequest")
    for c in commands:
        root.append(c)
    return ET.tostring(root, encoding="unicode")


def validate_file(path: str) -> list[str]:
    """Validate a hand written request XML file. Returns the command names it contains."""
    with open(path, encoding="utf-8") as fh:
        body = re.sub(r"<!--.*?-->", "", fh.read(), flags=re.S)
    root = ET.fromstring(body)
    if root.tag != "FabSuiteXMLRequest":
        raise PowerFabError(f"{path}: root is <{root.tag}>, expected <FabSuiteXMLRequest>.")
    names = []
    for child in root:
        if child.tag in ("ConnectionGUID", "Lang", "TrimbleIdentityToken",
                         "TrimbleIdentityAccessToken"):
            continue
        if COMMANDS and child.tag not in COMMANDS:
            raise PowerFabError(
                f"{path}: <{child.tag}> is not a command.{_suggest(child.tag, COMMANDS)}"
            )
        _validate_and_order(child, COMMANDS.get(child.tag, {}).get("fields"), child.tag)
        names.append(child.tag)
    return names


# -------------------------------------------------------------------- response parse


def _strip_ns(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


LB_PER_KG = 2.20462


def add_tons(rows: list[dict], weight_uom: str | None,
             columns=("Weight", "Total_Weight", "Received_Weight", "InStockWeight")) -> None:
    """
    Add a Tons column next to whichever weight column is present.

    Units differ BY COMMAND and are not negotiable on some of them.
    ProductionShippingProductivity_Get has no InputDisplayUnits and returns the system
    default, Imperial at MIW. PurchaseOrderSummary_Get has none either and has been seen
    returning METRIC in the same session, kg and mm. GetInventorySummary does accept
    InputDisplayUnits and defaults to metric if you leave it out. So never assume: read
    the WeightUOM the response echoes back and convert from that.
    """
    if not rows or not weight_uom:
        return
    # Two spellings arrive: InputDisplayUnits says Metric/Imperial, while a
    # fsDecimalWithUOM attribute says kg/lb outright. Accept either.
    u = str(weight_uom).strip().lower()
    if u in ("imperial", "lb", "lbs", "pound", "pounds"):
        f = 1.0 / 2000.0
    elif u in ("metric", "kg", "kgs", "kilogram", "kilograms"):
        f = LB_PER_KG / 2000.0
    else:
        return
    # Scan every row for the column, not just the first. Summary responses put a
    # grand-total row first and it often lacks the measure keys the detail rows carry,
    # so checking rows[0] alone silently skipped the whole conversion.
    present = {k for r in rows for k in r}
    for col in columns:
        if col in present:
            for r in rows:
                v = r.get(col)
                if isinstance(v, (int, float)):
                    r[col.replace("Weight", "Tons") if "Weight" in col else "Tons"] = round(v * f, 4)


def write_csv(rows: list[dict], path: str) -> int:
    """
    Write a list of dicts to CSV, with the union of all keys as the header.

    Rows from different pages or different group combinations do not always carry the
    same keys, so the columns are unioned in first-seen order rather than taken from
    row zero.
    """
    if not rows:
        open(path, "w").close()
        return 0
    cols: list[str] = []
    for r in rows:
        for k in r:
            if k not in cols:
                cols.append(k)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)
    return len(rows)


@dataclass
class Response:
    xml: str
    root: ET.Element = field(init=False)

    def __post_init__(self) -> None:
        self.root = ET.fromstring(self.xml)

    def errors(self) -> list[str]:
        return [e.text or "" for e in self.root.iter() if _strip_ns(e.tag) == "XMLError"]

    def warnings(self) -> list[str]:
        return [e.text or "" for e in self.root.iter() if _strip_ns(e.tag) == "XMLWarning"]

    def failures(self) -> list[str]:
        out = []
        for parent in self.root:
            for child in parent:
                if _strip_ns(child.tag) == "Successful" and (child.text or "").strip() != "1":
                    msg = next((c.text for c in parent
                                if _strip_ns(c.tag) == "ErrorMessage"), "(no message)")
                    out.append(f"{_strip_ns(parent.tag)}: {msg}")
        return out

    def raise_for_status(self, allow_warnings: bool = False) -> "Response":
        if self.errors():
            raise PowerFabError("XMLError (fatal, nothing ran): " + "; ".join(self.errors()))
        if self.warnings() and not allow_warnings:
            raise PowerFabError(
                "XMLWarning. PowerFab dropped an element it did not recognise and ran the "
                "command anyway, so this result is unfiltered or incomplete: "
                + "; ".join(self.warnings())
            )
        if self.failures():
            raise PowerFabError("; ".join(self.failures()))
        return self

    def commands(self, name: str) -> list[ET.Element]:
        """Every response block for a command name, in request order."""
        return [c for c in self.root if _strip_ns(c.tag) == name]

    def rows(self, name: str, index: int = 0, record_tag: str | None = None) -> list[dict]:
        """
        Flatten one command's repeating child records into dicts.

        PowerFab response shapes are barely documented, so rather than hard code a record
        element name this picks the most common repeating child unless you name one. Print
        the raw XML once for any new command and confirm, or read the response XSD.
        """
        blocks = self.commands(name)
        if not blocks:
            return []
        block = blocks[index]

        if record_tag is None:
            counts: dict[str, int] = {}
            for child in block:
                t = _strip_ns(child.tag)
                if t in ("Successful", "ErrorMessage", "CommandGUID"):
                    continue
                counts[t] = counts.get(t, 0) + 1
            if not counts:
                return []
            record_tag = max(counts, key=counts.get)

        rows = []
        for rec in block:
            if _strip_ns(rec.tag) != record_tag:
                continue
            row: dict[str, Any] = dict(rec.attrib)
            for f in rec.iter():
                if f is rec:
                    continue
                if len(f) == 0:
                    row[_strip_ns(f.tag)] = (f.text or "").strip()
                    row.update({f"{_strip_ns(f.tag)}@{k}": v for k, v in f.attrib.items()})
            rows.append(row)
        return rows

    # ---- the summary shape -------------------------------------------------
    #
    # ProductionShippingProductivity_Get and PurchaseOrderSummary_Get do not return a
    # flat record list. From fabsuite-xml-response-v0108.xsd their shape is:
    #
    #   <ResultData>
    #     <GetOptions>       an echo of what you asked for
    #     <InputDisplayUnits>  SizeUOM, LengthUOM, WeightUOM, ...
    #     <Result>           one per date bucket
    #       <StartDate/><EndDate/><Quantity/><Weight/><GrossWeight/><ModelWeight/>
    #       <Length/><Area/><SurfaceArea/><Hours/>
    #       <Group>          one per combination of the GroupBy values
    #         <Field FieldName="Station" DisplayName="Station">Saw 3</Field>
    #         <Field FieldName="Job #"   DisplayName="Job #">2381</Field>
    #         <FilterSet/>
    #         <Quantity/><Weight/><GrossWeight/><ModelWeight/><Length/><Area/>
    #         <SurfaceArea/><Hours/>
    #       </Group>
    #     </Result>
    #   </ResultData>
    #
    # summary_rows flattens that to one tidy row per date per group, which is what a
    # board or a pivot actually wants.

    _MEASURES = ("Quantity", "Weight", "GrossWeight", "ModelWeight",
                 "Length", "Area", "SurfaceArea", "Hours",
                 "POCount", "Cost", "AverageCost")

    def units(self, name: str, index: int = 0) -> dict:
        """InputDisplayUnits echoed back with the result. Tells you what Weight is in."""
        blocks = self.commands(name)
        if not blocks:
            return {}
        for u in blocks[index].iter():
            if _strip_ns(u.tag) == "InputDisplayUnits":
                return {_strip_ns(c.tag): (c.text or "").strip() for c in u}
        return {}

    def summary_rows(self, name: str, index: int = 0) -> list[dict]:
        blocks = self.commands(name)
        if not blocks:
            return []

        def num(el) -> dict:
            out = {}
            for c in el:
                t = _strip_ns(c.tag)
                if t in self._MEASURES and len(c) == 0:
                    v = (c.text or "").strip()
                    try:
                        out[t] = float(v) if v not in ("", None) else None
                    except ValueError:
                        out[t] = v
            return out

        def measures(el) -> dict:
            """
            Direct measures, plus any nested measure block, prefixed.

            ProductionShippingProductivity_Get puts Quantity/Weight/Hours straight on the
            Group. PurchaseOrderSummary_Get instead nests three blocks, Total, Received and
            Outstanding, each with POCount/Quantity/Weight/Length/Area/SurfaceArea/Cost.
            Flattening those to Received_Weight, Outstanding_Cost and so on keeps one row
            per date per group in both cases.
            """
            out = num(el)
            for c in el:
                tag = _strip_ns(c.tag)
                if tag in ("FilterSet", "Field") or len(c) == 0:
                    continue
                inner = num(c)
                if inner:
                    out.update({f"{tag}_{k}": v for k, v in inner.items()})
            return out

        rows: list[dict] = []
        for result in blocks[index].iter():
            if _strip_ns(result.tag) != "Result":
                continue
            dates = {_strip_ns(c.tag): (c.text or "").strip()
                     for c in result if _strip_ns(c.tag) in ("StartDate", "EndDate")}

            groups = [g for g in result if _strip_ns(g.tag) == "Group"]
            if not groups:
                # A date bucket with no GroupBy still carries its own totals.
                rows.append({**dates, **measures(result)})
                continue

            for g in groups:
                dims = {}
                for f in g:
                    if _strip_ns(f.tag) == "Field":
                        key = f.get("FieldName") or f.get("DisplayName") or "Field"
                        dims[key] = (f.text or "").strip()
                rows.append({**dates, **dims, **measures(g)})
        return rows

    def summary_to_csv(self, name: str, path: str, index: int = 0,
                       tons: bool = True) -> int:
        """
        Write summary_rows to CSV, adding a Tons column when the response says the
        weight came back Imperial. PowerFab's Imperial weight unit is pounds.
        """
        rows = self.summary_rows(name, index)
        if not rows:
            open(path, "w").close()
            return 0
        if tons and self.units(name, index).get("WeightUOM") == "Imperial":
            for r in rows:
                w = r.get("Weight")
                if isinstance(w, (int, float)):
                    r["Tons"] = round(w / 2000.0, 4)

        cols: list[str] = []
        for r in rows:
            for k in r:
                if k not in cols:
                    cols.append(k)
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=cols)
            w.writeheader()
            w.writerows(rows)
        return len(rows)

    # ---- inventory summary, its own shape again -----------------------------
    #
    #   <GetInventorySummary>
    #     <InputDisplayUnits/>  <SummaryField/>...
    #     <Summary>
    #       <SummaryField><Name>Location</Name><Value>..</Value><Display>..</Display></SummaryField>
    #       <InStockQuantity/> <InStockAmount UOM=".."/> <InStockWeight UOM=".."/>
    #       <OnOrderQuantity/> <OnOrderAmount UOM=".."/> <OnOrderWeight UOM=".."/>
    #     </Summary>
    #     <Offset/> <Limit/> <LimitHasMore/>
    #
    # Note this command pages. The server caps a response (5,000 rows observed) and sets
    # LimitHasMore, so a single call silently truncates unless you follow it.

    def inventory_rows(self, name: str = "GetInventorySummary", index: int = 0) -> list[dict]:
        blocks = self.commands(name)
        if not blocks:
            return []
        rows = []
        for summ in blocks[index]:
            if _strip_ns(summ.tag) != "Summary":
                continue
            row: dict[str, Any] = {}
            for c in summ:
                tag = _strip_ns(c.tag)
                if tag == "SummaryField":
                    key = next((f.text or "" for f in c if _strip_ns(f.tag) == "Name"), "Field")
                    val = next((f.text or "" for f in c if _strip_ns(f.tag) == "Display"), None)
                    if val is None:
                        val = next((f.text or "" for f in c if _strip_ns(f.tag) == "Value"), "")
                    row[key.strip()] = val.strip()
                else:
                    v = (c.text or "").strip()
                    try:
                        row[tag] = float(v) if v else 0.0
                    except ValueError:
                        row[tag] = v
                    if c.get("UOM"):
                        row[f"{tag}_UOM"] = c.get("UOM")
            rows.append(row)
        return rows

    def has_more(self, name: str = "GetInventorySummary", index: int = 0) -> bool:
        """True when the server truncated the response and there are more pages."""
        blocks = self.commands(name)
        if not blocks:
            return False
        for c in blocks[index]:
            if _strip_ns(c.tag) == "LimitHasMore":
                return (c.text or "").strip() in ("1", "true", "True")
        return False

    def to_csv(self, name: str, path: str, index: int = 0) -> int:
        rows = self.rows(name, index)
        if not rows:
            open(path, "w").close()
            return 0
        cols: list[str] = []
        for r in rows:
            for k in r:
                if k not in cols:
                    cols.append(k)
        with open(path, "w", newline="", encoding="utf-8") as fh:
            w = csv.DictWriter(fh, fieldnames=cols)
            w.writeheader()
            w.writerows(rows)
        return len(rows)


# ------------------------------------------------------------------------- the client


def _fix_exe(c):
    if not c.exe or not os.path.exists(c.exe):
        base = os.path.normpath(os.path.join(_HERE, ".."))   # the code folder
        for root, _dirs, files in os.walk(base):
            if "MIWFab.exe" in files:
                c.exe = os.path.join(root, "MIWFab.exe"); break
    return c


@dataclass
class Config:
    exe: str = r"C:\MIW\PowerFab\MIWFab.exe"
    mode: str = "remote"
    server: str = "192.168.1.15"
    port: int = 9154
    user: str = ""
    password: str = ""
    timeout: int = 300
    archive: str = ""      # folder to keep raw response XML in; "" disables

    @classmethod
    def load(cls, path: str = "powerfab.ini") -> "Config":
        c = cls()
        if not os.path.isabs(path) and not os.path.exists(path):
            path = os.path.join(_HERE, path)
        if not os.path.exists(path):
            return _fix_exe(c)
        p = configparser.ConfigParser()
        p.read(path)
        s = p["powerfab"]
        c.exe = s.get("exe", c.exe)
        c.mode = s.get("mode", c.mode)
        c.server = s.get("server", c.server)
        c.port = s.getint("port", c.port)
        c.user = s.get("user", c.user)
        c.password = s.get("password", c.password)
        c.timeout = s.getint("timeout", c.timeout)
        c.archive = s.get("archive", c.archive)
        # The ini's exe path is machine-specific (server = C:\Shares, work box = C:\Users).
        # If it does not exist on THIS machine, auto-find MIWFab.exe under the code folder.
        return _fix_exe(c)


class PowerFab:
    def __init__(self, config: Config | None = None) -> None:
        self.cfg = config or Config.load()

    def send(self, request_xml: str) -> Response:
        d = tempfile.mkdtemp(prefix="miwfab_")
        req_path = os.path.join(d, "request.xml")
        res_path = os.path.join(d, "response.xml")
        with open(req_path, "w", encoding="utf-8") as fh:
            fh.write(request_xml)

        args = [self.cfg.exe,
                "--mode", self.cfg.mode,
                "--server", self.cfg.server,
                "--port", str(self.cfg.port),
                "--user", self.cfg.user,
                "--pass", self.cfg.password,
                "--in", req_path,
                "--out", res_path]
        proc = subprocess.run(args, capture_output=True, text=True, timeout=self.cfg.timeout)

        if not os.path.exists(res_path):
            raise PowerFabError(
                f"MIWFab.exe produced no response (exit {proc.returncode}).\n{proc.stderr.strip()}")
        with open(res_path, encoding="utf-8") as fh:
            xml = fh.read()

        # Keep the raw response. Every shape surprise so far (nested measures, the 5,000
        # row cap, metric units) was diagnosed from a saved response, and re-running a
        # pull to see the XML again needs the MIW network.
        if self.cfg.archive:
            try:
                os.makedirs(self.cfg.archive, exist_ok=True)
                import datetime as _dt
                stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
                first = ET.fromstring(xml)
                label = next((_strip_ns(c.tag) for c in first
                              if _strip_ns(c.tag) not in ("Lang", "XMLWarning", "XMLError")),
                             "response")
                with open(os.path.join(self.cfg.archive, f"{stamp}_{label}.xml"),
                          "w", encoding="utf-8") as ah:
                    ah.write(xml)
            except Exception:
                pass   # archiving is a convenience, never a reason to fail a pull

        return Response(xml)

    def run(self, *commands: ET.Element, allow_warnings: bool = False) -> Response:
        return self.send(build_request(*commands)).raise_for_status(allow_warnings)

    def version(self) -> dict:
        r = self.send(build_request(ET.Element("Version"))).raise_for_status()
        blocks = r.commands("Version")
        if not blocks:
            raise PowerFabError("No Version block in response.")
        return {_strip_ns(c.tag): (c.text or "").strip() for c in blocks[0]}


# ---------------------------------------------------------------------------- helpers


def filters(*specs: tuple) -> dict:
    """
    Build a Filters block.

    PowerFab filter logic is fixed and cannot be changed:
        same FilterType repeated  -> OR
        different FilterType      -> AND
    So ("Job","2381"), ("Job","2382"), ("Sequence","5") means
        Job IN (2381, 2382) AND Sequence = 5
    """
    out: list[dict] = []
    for spec in specs:
        ftype, value = spec[0], spec[1]
        match = spec[2] if len(spec) > 2 else None
        f: dict[str, Any] = {"FilterType": ftype, "FilterValue": {"Value": value}}
        if match:
            f["MatchType"] = match
        out.append(f)
    return {"Filter": out}


IMPERIAL = {
    "SizeUOM": "Imperial",
    "LengthUOM": "Imperial",
    "LengthTypeIDMetric": 3,      # millimeters, unused when LengthUOM is Imperial
    "LengthTypeIDImperial": 4,    # decimal feet, the easiest to do arithmetic on
    "WeightUOM": "Imperial",      # pounds
    "PriceUOM": "Imperial",
}
"""
InputDisplayUnits for a request that should come back in pounds and feet.

All six fields are REQUIRED by the schema, so you cannot send a partial block. Commands
that accept InputDisplayUnits and are not given it default to METRIC: the first inventory
pull came back with weights in kg and lengths in mm, silently. ProductionShippingProductivity_Get
has no InputDisplayUnits at all and returns the system default, which at MIW is Imperial,
so the two disagree unless you are explicit.
"""


def date_window(days_back: int = 14, today: str | None = None) -> tuple[str, str]:
    """
    Trailing window for the nightly pull.

    Fourteen days, re-pulled and overwritten every night, because backdated scans and
    rework land on prior dates and the productivity call has no LastLogID.
    """
    import datetime as _dt
    end = _dt.date.fromisoformat(today) if today else _dt.date.today()
    return (end - _dt.timedelta(days=days_back)).isoformat(), end.isoformat()


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        print(describe(sys.argv[1]))
    else:
        print(f"{len(COMMANDS)} commands loaded from {SCHEMA.get('source', '(no schema)')}")
        print("Usage: python powerfab.py <CommandElementName>   to see what it accepts")
