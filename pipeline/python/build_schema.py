"""
build_schema.py
Turn Tekla's shipped XSD into a compact JSON catalog that powerfab.py validates against.

Run this once, and again after any PowerFab upgrade:

    python build_schema.py "C:\\Program Files\\Tekla\\Tekla PowerFab\\Docs\\fabsuite-xml-request-v0108.xsd"

It writes powerfab_schema.json beside this file.

Why this exists. The developer website is a guide, not a spec, and it was wrong about
three separate things in one morning: class names are not element names, ConnectionGUID
is an element and not an attribute, and the enumeration members it publishes are not the
strings the schema accepts (it lists productivityproductionstation, the schema says
productivity-production-station). The XSD that ships with PowerFab is the only authority.
Validating against it offline turns a round trip and a cryptic rejection into an instant
local error with the legal values printed.
"""

from __future__ import annotations

import json
import os
import sys
import xml.etree.ElementTree as ET

XSD = "{http://www.w3.org/2001/XMLSchema}"


def q(tag: str) -> str:
    return XSD + tag


def main(path: str) -> int:
    tree = ET.parse(path)
    root = tree.getroot()

    # Named complexTypes and simpleTypes, so we can follow type= references.
    complex_types = {c.get("name"): c for c in root.findall(q("complexType")) if c.get("name")}
    simple_types = {s.get("name"): s for s in root.findall(q("simpleType")) if s.get("name")}

    def enum_values(node) -> list[str] | None:
        """Inline simpleType restriction, or a named simpleType reference."""
        st = node.find(q("simpleType"))
        if st is None:
            tname = node.get("type")
            if tname in simple_types:
                st = simple_types[tname]
        if st is None:
            return None
        vals = [e.get("value") for e in st.iter(q("enumeration"))]
        return vals or None

    def fields_of(ctype, depth: int = 0) -> list[dict]:
        """Direct child elements of a complexType, following complexContent extensions."""
        if ctype is None or depth > 6:
            return []
        out: list[dict] = []

        # An extension inherits its base type's fields first (fsreqCommand gives every
        # command its CommandGUID and APILog).
        for ext in ctype.iter(q("extension")):
            base = ext.get("base")
            if base in complex_types and base != ctype.get("name"):
                out.extend(fields_of(complex_types[base], depth + 1))
            break

        for el in ctype.iter(q("element")):
            name = el.get("name")
            if not name:
                continue
            if any(f["name"] == name for f in out):
                continue
            f = {
                "name": name,
                "type": el.get("type"),
                "required": el.get("minOccurs", "1") != "0",
                "repeatable": el.get("maxOccurs", "1") not in ("1", None),
            }
            ev = enum_values(el)
            if ev:
                f["enum"] = ev
            doc = el.find(f"{q('annotation')}/{q('documentation')}")
            if doc is not None and (doc.text or "").strip():
                f["doc"] = " ".join(doc.text.split())
            out.append(f)
        return out

    # The request root: its child elements are the legal command element names.
    req = None
    for e in root.findall(q("element")):
        if e.get("name") == "FabSuiteXMLRequest":
            req = e
            break
    if req is None:
        print("Could not find the FabSuiteXMLRequest root element.", file=sys.stderr)
        return 2

    inline = req.find(q("complexType"))
    ctype = inline if inline is not None else complex_types.get(req.get("type"))

    commands: dict[str, dict] = {}
    for el in ctype.iter(q("element")):
        name, tname = el.get("name"), el.get("type")
        if not name:
            continue
        entry: dict = {"type": tname}
        doc = el.find(f"{q('annotation')}/{q('documentation')}")
        if doc is not None and (doc.text or "").strip():
            entry["doc"] = " ".join(doc.text.split())
        if tname in complex_types:
            entry["fields"] = fields_of(complex_types[tname])
        else:
            ev = enum_values(el)
            if ev:
                entry["enum"] = ev
        commands[name] = entry

    # Child complexTypes referenced by command fields, one level down, which is where
    # every GetOptions block lives.
    nested: dict[str, list[dict]] = {}
    for entry in commands.values():
        for f in entry.get("fields", []):
            t = f.get("type")
            if t in complex_types and t not in nested:
                nested[t] = fields_of(complex_types[t])

    out = {
        "source": os.path.basename(path),
        "commands": commands,
        "types": nested,
    }
    dest = os.path.join(os.path.dirname(os.path.abspath(__file__)), "powerfab_schema.json")
    with open(dest, "w", encoding="utf-8") as fh:
        json.dump(out, fh, indent=1, sort_keys=True)

    n_enum = sum(1 for e in commands.values() for f in e.get("fields", []) if "enum" in f)
    print(f"{len(commands)} commands, {len(nested)} nested types, "
          f"{n_enum} enumerated fields -> {dest}")
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    sys.exit(main(sys.argv[1]))
