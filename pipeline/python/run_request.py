"""Run one request XML through MIWFab.exe and save the raw response.
Usage:  py run_request.py <request.xml> <response.xml>
Credentials come from powerfab.ini, same as the nightly pull."""
import sys, os, subprocess, powerfab
cfg = powerfab.Config.load()
req = os.path.abspath(sys.argv[1]); out = os.path.abspath(sys.argv[2])
args = [cfg.exe, "--mode", cfg.mode, "--server", cfg.server, "--port", str(cfg.port),
        "--user", cfg.user, "--pass", cfg.password, "--in", req, "--out", out]
print("running:", os.path.basename(req))
p = subprocess.run(args, capture_output=True, text=True, timeout=cfg.timeout)
print("exit", p.returncode)
if p.stdout.strip(): print(p.stdout.strip())
if p.stderr.strip(): print("ERR:", p.stderr.strip())
print("wrote", out if os.path.exists(out) else "(NO OUTPUT FILE)")
