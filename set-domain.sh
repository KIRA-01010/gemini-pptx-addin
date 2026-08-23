#!/usr/bin/env bash
# Usage: ./set-domain.sh yourname.github.io/gemini-pptx-addin
# Rewrites every https://YOUR-DOMAIN/... URL in manifest.xml to your real
# hosting domain (no protocol, no trailing slash). Run this once after you've
# hosted the files, before sideloading manifest.xml into PowerPoint.

set -euo pipefail

if [ -z "${1:-}" ]; then
  echo "Usage: ./set-domain.sh <domain-and-optional-path>"
  echo "Example: ./set-domain.sh yourname.github.io/gemini-pptx-addin"
  exit 1
fi

DOMAIN="${1%/}"  # strip trailing slash if present

sed -i.bak "s#https://YOUR-DOMAIN#https://${DOMAIN}#g" manifest.xml
rm -f manifest.xml.bak

echo "manifest.xml now points at https://${DOMAIN}"
echo "Next: upload/push the folder contents to that host, then sideload manifest.xml in PowerPoint."
