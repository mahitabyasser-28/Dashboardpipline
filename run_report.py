"""
Runs pipeline_health.py on GitHub Actions.

Logs in to Salesforce with the credentials stored as GitHub secrets,
hands the connection to pipeline_health.py as `sf` (exactly like the
notebook does), then copies the finished report to site/index.html
so GitHub Pages can serve it.

pipeline_health.py itself is NOT changed: paste your script in as-is.
"""
import glob, os, runpy, shutil, sys

from simple_salesforce import Salesforce

missing = [k for k in ("SF_USERNAME", "SF_PASSWORD", "SF_TOKEN") if not os.environ.get(k)]
if missing:
    sys.exit(f"❌ missing GitHub secret(s): {', '.join(missing)} "
             "(Settings → Secrets and variables → Actions)")

sf = Salesforce(
    username=os.environ["SF_USERNAME"],
    password=os.environ["SF_PASSWORD"],
    security_token=os.environ["SF_TOKEN"],
    domain=os.environ.get("SF_DOMAIN") or "login",   # set SF_DOMAIN=test only for a sandbox
)
print("✅ Salesforce connected:", sf.sf_instance)

# Clear old reports so we always publish the one from this run
for old in glob.glob("Open_Pipeline_Health_*.html"):
    os.remove(old)

# Run the script with `sf` already defined, the same as in the notebook
runpy.run_path("pipeline_health.py", init_globals={"sf": sf}, run_name="__main__")

reports = glob.glob("Open_Pipeline_Health_*.html")
if not reports:
    sys.exit("❌ pipeline_health.py finished but wrote no report; see the log above")

os.makedirs("site", exist_ok=True)
shutil.copy(reports[0], "site/index.html")
with open("site/robots.txt", "w") as f:          # keep the page out of Google
    f.write("User-agent: *\nDisallow: /\n")
print(f"🌐 {reports[0]} → site/index.html")
