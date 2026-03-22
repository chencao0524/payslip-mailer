# 工资条发送器

This repository contains the desktop-only 工资条发送器 tool. It can be run or packaged on macOS and Windows as an independent desktop app.

## Stack
- UI: React 18 rendered through Vite
- Shell: Tauri (Rust) providing native bundling and filesystem access
- Runtime: Node.js toolchain for development, Rust toolchain for bundling

## Running locally
1. Install the Node.js toolchain (Node 20+) and Rust (stable). `npm install` will populate `node_modules`.
2. Run `npm run dev` from this directory. That launches the Vite dev server and Tauri devwrapper in a single command.
3. The UI issues `invoke` calls to the Tauri backend, which appends a history file in `app_config_dir()/payslip-mailer/history.json`.

> If `npm install` hangs because the sandbox cannot reach the public registry, download the dependencies elsewhere or mirror them into `node_modules` so the dev server can start.

## Packaging for macOS and Windows
- `npm run build` -> Tauri build. By default it already targets `dmg` and `msi`. Specify `TAURI_BUNDLE_TARGET` if you need additional CPU architectures (`x86_64-apple-darwin`, `x86_64-pc-windows-msvc`, etc.).
- Tauri was chosen over Electron because it ships a much smaller binary, reuses the system WebView, includes Rust-based API access, and has first-class support for macOS & Windows bundles. Electron can be added later if a Chromium runtime is explicitly required, but Tauri keeps the tooling lightweight for this focused use case.

## macOS signing and notarization
- The repository is now set up for hardened-runtime macOS bundles with [src-tauri/Entitlements.plist](/Users/chencao/Developer/payslip-mailer/src-tauri/Entitlements.plist) and `bundle.macOS` settings in [src-tauri/tauri.conf.json](/Users/chencao/Developer/payslip-mailer/src-tauri/tauri.conf.json).
- The app identifier is `cn.beeworks.payslip`. Use the same identifier in your Apple Developer signing setup.
- Before packaging for other Macs, check the local machine first: `npm run check:macos:signing`
- Local packaging without signing: `npm run build:macos:app`
- Architecture-specific packaging:
- `npm run build:macos:dmg:arm64` for Apple Silicon Macs
- `npm run build:macos:dmg:x64` for Intel Macs
- `npm run build:windows:msi` on a Windows host for Windows installers
- `npm run build:release` builds all bundles supported by the current host:
- on macOS: `arm64` and `x64` dmg bundles
- on Windows: `x64` msi bundle
- Set `NOTARIZE_MACOS_DMG=1` when running the release or macOS build scripts if you want the generated ASCII dmg copies to be notarized automatically.
- If you prefer not to export the base64 certificate manually, set `APPLE_CERTIFICATE_FILE` to your `.p12` path and `APPLE_CERTIFICATE_PASSWORD` to its password before running the macOS build scripts.
- If you want one build per CPU family, distribute the matching dmg. The current scripts do not produce a universal macOS binary.
- The build script also creates ASCII-named dmg copies such as `src-tauri/target/release/bundle/dmg/payslip-mailer-arm64.dmg` because `stapler` on this machine does not reliably open the Chinese dmg filename.
- Signed app or dmg packaging can use either a local Keychain certificate or exported Apple credentials. For a local Keychain workflow, install a valid `Developer ID Application` certificate in Keychain Access so `security find-identity -v -p codesigning` returns a usable identity.
- Tauri 2 can also consume signing credentials from environment variables:
- `APPLE_CERTIFICATE`: base64 encoded `.p12` signing certificate
- `APPLE_CERTIFICATE_PASSWORD`: password for that `.p12`
- `APPLE_API_KEY`: App Store Connect API key ID
- `APPLE_API_ISSUER`: App Store Connect issuer ID
- `APPLE_API_KEY_PATH`: path to the `.p8` private key file
- After those are ready, run `npm run build:macos:dmg:arm64` or `npm run build:macos:dmg:x64`.
- If `stapler` fails on the generated Chinese-named dmg, run `npm run notarize:macos:dmg -- "<path-to-dmg>"`. This script notarizes and staples an ASCII-named copy of the dmg because `stapler` on this machine does not reliably open the Chinese file name.
- To let other Macs install the app without the usual Gatekeeper failure path, you should do both:
- sign with `Developer ID Application`
- notarize with the App Store Connect API key variables above
- If you skip signing or notarization, users may see the app blocked, or double-clicking may appear to do nothing until they manually approve it in System Settings.

## Notes
- The app writes history to the standard config directory via `tauri::api::path::app_config_dir` and keeps the UI self-contained.
- You can tweak the recipient list, subject, and note fields before hitting the send button; the frontend simply passes them through to Rust for logging.

## 截图
<img width="283" height="222" alt="图片" src="https://github.com/user-attachments/assets/d743c3a5-a642-4b83-8640-5ff40c77cb57" />

模版：[工资条模版.xlsx](https://github.com/user-attachments/files/26165805/default.xlsx)


