#!/bin/zsh

set -euo pipefail

echo "Checking macOS code signing identities in the current keychains..."
security find-identity -v -p codesigning

echo
echo "Looking for Developer ID Application certificates..."
developer_id_app_identities=$(
  security find-certificate -a -c "Developer ID Application" -Z 2>/dev/null | \
    awk '/SHA-256 hash:|labl:/' || true
)

if [[ -n "${developer_id_app_identities}" ]]; then
  echo "${developer_id_app_identities}"
else
  echo "No Developer ID Application certificate found in the accessible keychains."
fi

echo
missing=()
for var_name in APPLE_API_KEY APPLE_API_ISSUER APPLE_API_KEY_PATH; do
  if [[ -z "${(P)var_name:-}" ]]; then
    missing+=("${var_name}")
  fi
done

if [[ ${#missing[@]} -eq 0 ]]; then
  echo "Notarization environment variables are present."
else
  echo "Missing notarization environment variables: ${missing[*]}"
fi

echo
echo "Expected setup:"
echo "- Signing requires a valid 'Developer ID Application' identity in Keychain Access."
echo "- Notarization requires APPLE_API_KEY, APPLE_API_ISSUER, and APPLE_API_KEY_PATH."
