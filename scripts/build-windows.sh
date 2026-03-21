#!/bin/zsh

set -euo pipefail

arch="${1:-x64}"
bundle="${2:-msi}"

case "${arch}" in
  x64)
    target="x86_64-pc-windows-msvc"
    ;;
  *)
    echo "Unsupported Windows architecture: ${arch}" >&2
    echo "Use one of: x64" >&2
    exit 1
    ;;
esac

case "${bundle}" in
  msi)
    ;;
  *)
    echo "Unsupported Windows bundle type: ${bundle}" >&2
    echo "Use one of: msi" >&2
    exit 1
    ;;
esac

host_os="$(uname -s)"
case "${host_os}" in
  MINGW*|MSYS*|CYGWIN*|Windows_NT)
    ;;
  *)
    echo "Windows MSI packaging must be run on a Windows host." >&2
    echo "Current host: ${host_os}" >&2
    echo "Run this script in the project directory on Windows:" >&2
    echo "  ./scripts/build-windows.sh ${arch} ${bundle}" >&2
    exit 2
    ;;
esac

echo "Building Windows ${bundle} for ${arch} (${target})..."
export TAURI_BUNDLE_TARGET="${target}"
node node_modules/@tauri-apps/cli/tauri.js build --bundles "${bundle}"

