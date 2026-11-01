# Creator's Guide: Build, Package, and Release LAYA

This guide is for the maintainer who builds and publishes LAYA Agentic Studio.
It covers local packages for Windows, macOS, and Linux, plus the automated
GitHub release workflow and the signed-update design.

> **Important:** Build packages on the operating system they target. PyInstaller
> does not make a Windows build on macOS or a Linux build on Windows. The
> repository's GitHub Actions workflow is the supported way to build all
> configured platforms from one release tag.

## 1. What the package contains

The `packaging/build_release.py` script uses PyInstaller to bundle the Python
application and its Python dependencies. It also builds the update-helper
executable. The installer does **not** include:

- An Ollama server or model weights. Users configure a local/remote Ollama
  server or an OpenAI-compatible endpoint in the app.
- Tesseract OCR. Install it separately if image/scanned-document OCR is
  needed.
- Microsoft Word, LibreOffice, or other desktop software needed for optional
  legacy document conversion.

Before bundling, the release builder checks every `CTkScrollableFrame` call
against the `customtkinter` version installed in that build environment. If
the app uses an unsupported constructor argument, the build stops with the
offending source line instead of producing an executable that crashes at
startup. The builder also clears its named output folder and platform release
artifacts before rebuilding, preventing an old executable from being merged
into a new package. Do not distribute manually copied or previously generated
files from `dist/`; use artifacts from the same successful build run.

Treat the build as a distributable application, not as a way to make Python
source impossible to inspect. Do not include credentials, signing private
keys, personal documents, databases, or model weights in a package.

## 2. Prepare a clean release checkout

1. Make sure the intended changes are committed and pushed to the repository.
2. Start from a clean checkout of the exact commit you intend to release.
3. Check `git status --short` and confirm no local settings, databases,
   documents, secrets, or generated files will enter the release.
4. Choose a new semantic version, for example `0.2.0`. Published versions
   should increase monotonically; do not reuse an existing release tag.
5. Use Python 3.13 for consistency with the current GitHub Actions workflow.

The builder writes the selected version into `app_version.py` in its checkout.
On GitHub Actions this change exists only in the temporary runner. For a local
build, use a disposable clean worktree or restore the generated version change
afterward; do not discard other work in a dirty checkout.

## 3. Build on Windows

Open PowerShell in the repository root:

```powershell
py -3.13 -m venv .venv-build
.\.venv-build\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt -r packaging/requirements-build.txt
```

Install **Inno Setup 6** to create the per-user Windows installer. It must be
available as `ISCC.exe` on `PATH` or in its standard Program Files location.
Then build:

```powershell
python packaging/build_release.py 0.2.0
```

For a portable ZIP smoke build without Inno Setup:

```powershell
python packaging/build_release.py 0.2.0 --portable-only
```

Expected output in `dist\`:

- `LAYA-Windows-Setup.exe` — per-user installer; no administrator privileges
  should be required.
- `LAYA-Windows-x64.zip` — portable/update archive.

Test the installer on a clean Windows user account or disposable VM. Verify
launch, first-run settings, model-provider configuration, document import and
export, voice behavior where supported, and uninstall. Ollama and Tesseract
must be tested as separately installed components.

## 4. Build on macOS

Use a Mac with Xcode Command Line Tools and Python 3.13. Install the native
audio build dependency, then install the Python requirements:

```bash
xcode-select --install
brew install portaudio
python3.13 -m venv .venv-build
source .venv-build/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt -r packaging/requirements-build.txt
python packaging/build_release.py 0.2.0
```

Expected output in `dist/`:

- `LAYA-macOS-<architecture>.dmg` — disk image containing the app bundle and
  per-user install command.
- `LAYA-macOS-<architecture>.zip` — app bundle/update archive.

The architecture is determined by the build machine (`x64` or `arm64`). Build
on each architecture you intend to support and verify the corresponding
package is present in the release manifest. A universal macOS binary is not
currently produced.

To install from the DMG, open it and run `install_macos.command` from the
mounted image. It copies the app to `~/Applications`. The current release
scripts do not sign or notarize macOS packages, so macOS Gatekeeper may warn
or block launch. For public distribution, configure an Apple Developer
Developer ID signing identity, notarize the app/DMG, and test the stapled
package on a clean Mac.

## 5. Build on Linux

Build on a Linux distribution compatible with the oldest distribution you
intend to support. The current CI uses Ubuntu 24.04; building on an older
baseline can improve compatibility with older glibc systems, but that baseline
must also be tested.

On Ubuntu/Debian, install the native build dependencies:

```bash
sudo apt-get update
sudo apt-get install -y python3.13 python3.13-venv portaudio19-dev python3-tk
python3.13 -m venv .venv-build
source .venv-build/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt -r packaging/requirements-build.txt
python packaging/build_release.py 0.2.0
```

Expected output:

- `LAYA-Linux-x64.tar.gz` — application bundle with `install_linux.sh`.

Extract the archive, then install from the extracted directory:

```bash
mkdir -p ~/Downloads/laya-release
tar -xzf LAYA-Linux-x64.tar.gz -C ~/Downloads/laya-release
cd ~/Downloads/laya-release
bash install_linux.sh
```

The installer places the app under the user's XDG data directory, creates a
launcher under `~/.local/bin`, and adds a desktop entry. Confirm that
`~/.local/bin` is on the user's `PATH`. Test on a clean supported Linux
account, including audio/Tk availability and any required shared system
libraries.

## 6. Build all configured platforms with GitHub Actions

The workflow at `.github/workflows/release.yml` builds each platform on its
native GitHub-hosted runner, uploads the build artifacts, signs an update
manifest, and publishes a GitHub Release.

Before tagging:

1. Run the application and packaging checks locally.
2. Verify each supported platform's artifact names and architectures.
3. Confirm release-signing prerequisites in [OTA signing and update readiness](#8-ota-signing-and-update-readiness).
4. Ensure the release commit is on the intended branch.

Then create and push a new tag:

```bash
git tag v0.2.0
git push origin v0.2.0
```

Alternatively, run **Build signed desktop releases** from the repository's
GitHub Actions page with `workflow_dispatch`, supplying an existing tag. Do not
use a tag for a commit that has not passed review.

The workflow uploads platform artifacts for seven days, then publishes them
to the GitHub Release. The release includes the installers/archives,
`update-manifest.json`, and `update-manifest.sig`. Inspect the workflow run and
release page; a green build alone is not a substitute for installing and
smoke-testing each package.

Current workflow targets:

| Platform | Runner | Intended release assets |
|---|---|---|
| Windows x64 | `windows-latest` | Setup installer and x64 ZIP |
| macOS runner architecture | `macos-15` | DMG and ZIP for the runner architecture |
| Linux x64 | `ubuntu-24.04` | x64 tarball |

This matrix does not currently promise Windows ARM64, Linux ARM64, or a
universal macOS package. Add and test explicit runners/build targets before
advertising those architectures.

## 7. Install and smoke-test release artifacts

Test each artifact on a clean machine or VM, not only on the build machine.
Keep a copy of the previous release available for rollback.

- **Windows:** run `LAYA-Windows-Setup.exe`, launch from the Start menu, and
  verify uninstall. The ZIP can be extracted and run as a portable build.
- **macOS:** open the DMG and run the included install command. Confirm the
  app is copied to `~/Applications`; test Gatekeeper behavior on a clean Mac.
- **Linux:** extract the tarball and run `bash install_linux.sh`; verify the
  launcher and desktop entry.

For each platform, test startup without a source checkout or developer
virtualenv. Then check provider discovery, a small model request, document
extraction, each supported export format, settings persistence, and update
availability. Test optional features only when their external service or
system dependency is installed.

## 8. OTA signing and update readiness

The intended update design is:

1. The release workflow hashes each platform update archive.
2. It signs `update-manifest.json` using an Ed25519 private key stored as the
   GitHub Actions secret `LAYA_UPDATE_SIGNING_KEY_B64`.
3. The app verifies the signature using the public key embedded in
   `update_service.py`, checks the archive hash, downloads the matching
   package, then starts the update helper after the app exits.

The private key must never be committed, uploaded as a release asset, or
embedded in the application. Back it up in a secure password manager or
hardware-backed secret store. Losing it requires a deliberate application
update that changes the trusted public key before future signed updates can
work.

### Critical checks before enabling OTA for users

Do not announce OTA as production-ready until all of these are fixed and
verified end to end:

1. **Signing-key match:** `update_service.py` contains a pinned public key.
   The key generated by `tools/prepare_update_signing_key.py` must be the
   private-key counterpart of that exact public key. The helper script prints
   a public key but does not update the application. Verify the match in a
   controlled test; never replace the public key without coordinating the
   corresponding GitHub secret.
2. **Install/update layout:** test the exact archive generated for each OS
   against the installed path used by that OS's updater. Confirm replacement,
   rollback on failure, and relaunch.
3. **Authenticode and Apple notarization:** the signed update manifest
   authenticates the update archive but does not code-sign the Windows
   installer or notarize macOS software. Configure Windows code signing and
   Apple Developer ID signing/notarization for production distribution.

After these prerequisites are satisfied:

1. Run `python tools/prepare_update_signing_key.py` in a trusted maintainer
   environment. The tool stores the PEM in the OS-specific per-user
   configuration directory and prints a raw private key and its public key.
   Treat the private value as a secret; do not put it in a ticket, source file,
   or chat.
2. Replace `UPDATE_PUBLIC_KEY_B64` in `update_service.py` with the printed
   public key and commit that public-key change before publishing an app that
   trusts the new signer.
3. Add the printed base64 raw private key as the GitHub Actions secret
   `LAYA_UPDATE_SIGNING_KEY_B64`. Back up the PEM securely.
4. Publish a higher version tag using the workflow above.
5. Verify the release manifest signature and each listed SHA-256 digest.
6. Test **Help → Check for Updates** from an older installed build on each
   target OS, including cancel, download failure, invalid signature, and
   successful installation.

OTA packages replace the complete application bundle; they are not binary
delta patches. User data is intended to remain in the OS-specific per-user
data directory and must be checked across upgrades.

## 9. Patch and maintenance release procedure

For a normal code patch:

1. Make the change on a branch and add/update tests.
2. Run the smallest relevant test set, compile the application, and build the
   package on the target OS.
3. Verify that app data, settings, and existing databases remain compatible.
4. Review the diff for secrets, local databases, generated files, and
   accidental model/document inclusion.
5. Merge the reviewed change, choose a new version, and publish its tag.
6. Install the release on a clean machine and keep the previous package until
   the new one is verified.

Use a patch version for compatible fixes (for example `0.2.0` to `0.2.1`).
Use a minor or major version when the change materially alters features or
compatibility. Never overwrite an existing published release tag.

## 10. Troubleshooting

- **PyInstaller fails to import a package:** add the necessary hidden import or
  data collection in `packaging/build_release.py`, rebuild in a clean
  environment, and test the installed package rather than only the source app.
- **Audio build fails:** install the platform's PortAudio development/runtime
  package and confirm the audio device and permission settings.
- **App starts but OCR does not work:** install Tesseract separately and
  verify its executable is discoverable by the app.
- **No models appear:** check the configured provider endpoint, network,
  authentication, and whether the server is running. Ollama/model weights are
  not bundled.
- **Linux package fails on an older distribution:** rebuild on an older
  compatible baseline and test all supported distributions.
- **macOS blocks launch:** the current artifacts are not notarized. Do not ask
  users to bypass security controls for a public release; sign and notarize
  the package.
- **OTA check fails:** inspect the signed manifest assets, repository/release
  version, public/private key match, archive hash, platform identifier, and
  update-helper filename. Do not disable signature or hash validation to
  make an update succeed.

## 11. Release checklist

- [ ] Clean reviewed commit and new version tag.
- [ ] Python/runtime dependencies install from the checked-in manifests.
- [ ] Native build dependencies installed on each platform runner.
- [ ] Build artifacts contain no secrets, databases, documents, or model
      weights.
- [ ] Installer tested on a clean account/VM for each supported OS.
- [ ] External dependencies (Ollama, Tesseract, audio permissions) explained
      to users.
- [ ] Windows signing and macOS signing/notarization addressed for public
      distribution.
- [ ] OTA signing key matches the embedded public key.
- [ ] Update helper names and install layouts match and full OTA tests pass.
- [ ] GitHub Release assets and manifest hashes verified.
- [ ] Previous release retained for rollback.
