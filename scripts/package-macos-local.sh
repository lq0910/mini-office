#!/usr/bin/env bash
# 功能：本机生成签名应用、标准 DMG、ZIP 和校验清单，避免依赖 Finder AppleScript；时间：2026-11-11；作者：lq。
set -euo pipefail
cd "$(dirname "$0")/.."
if [[ "${MINI_OFFICE_SKIP_NATIVE_BUILD:-0}" != "1" ]]; then
  APPLE_SIGNING_IDENTITY="${APPLE_SIGNING_IDENTITY:--}" pnpm tauri:build --bundles app --ci
fi
app_path="target/release/bundle/macos/Mini Office.app"
release_version=$(/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' "$app_path/Contents/Info.plist")
release_arch=$(lipo -archs "$app_path/Contents/MacOS/mini-office-desktop")
case "$release_arch" in
  arm64) release_suffix=arm64 ;;
  x86_64) release_suffix=x64 ;;
  'x86_64 arm64'|'arm64 x86_64') release_suffix=universal ;;
  *) echo "不支持的架构：$release_arch" >&2; exit 1 ;;
esac
release_root="$PWD/output/releases/$release_version"
mkdir -p "$release_root"
staging_root=$(mktemp -d "$release_root/.dmg-root.XXXXXX")
# 功能：只清理由本脚本创建的临时镜像目录；时间：2026-11-11；作者：lq。
trap 'rm -rf "$staging_root"' EXIT
codesign --verify --deep --strict "$app_path"
ditto "$app_path" "$staging_root/Mini Office.app"
ln -s /Applications "$staging_root/Applications"
release_name="Mini-Office_${release_version}_macos-${release_suffix}"
hdiutil create -ov -volname "Mini Office $release_version" -srcfolder "$staging_root" -format UDZO -o "$release_root/$release_name.dmg"
ditto -c -k --sequesterRsrc --keepParent "$app_path" "$release_root/$release_name.zip"
hdiutil verify "$release_root/$release_name.dmg"
cd "$release_root"
shasum -a 256 "$release_name.dmg" "$release_name.zip" > "SHA256SUMS-macos-$release_suffix.txt"
shasum -a 256 -c "SHA256SUMS-macos-$release_suffix.txt"
