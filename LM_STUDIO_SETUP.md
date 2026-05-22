# LM Studio Setup

[LM Studio](https://lmstudio.ai/) runs local LLMs and exposes an
OpenAI-compatible HTTP server on `http://localhost:1234/v1`.

## Install

### Windows (PowerShell)

```powershell
irm https://lmstudio.ai/install.ps1 | iex
```

`irm` is `Invoke-RestMethod` and `iex` is `Invoke-Expression`. The
command downloads the official installer script and runs it. Inspect the
script first if you'd rather not pipe a remote payload straight into
your shell:

```powershell
irm https://lmstudio.ai/install.ps1 | Out-File -FilePath install.ps1
# review install.ps1, then:
./install.ps1
```

### macOS

Download the `.dmg` from <https://lmstudio.ai/> or use Homebrew:

```sh
brew install --cask lm-studio
```

### Linux

Download the AppImage from <https://lmstudio.ai/>:

```sh
chmod +x LM-Studio-*.AppImage
./LM-Studio-*.AppImage
```

## CLI (`lms`)

After installing the desktop app, bootstrap the CLI:

```sh
# macOS / Linux
~/.lmstudio/bin/lms bootstrap

# Windows
%USERPROFILE%\.lmstudio\bin\lms.exe bootstrap
```

Common commands:

```sh
lms ls                       # list installed models
lms get <model>              # download a model
lms server start --port 1234 # start the OpenAI-compatible server
lms server status
```

## Using the local server

Point any OpenAI-compatible client at `http://localhost:1234/v1` with a
placeholder API key:

```sh
curl http://localhost:1234/v1/chat/completions \
  -H "Content-Type: application/json" \
  -d '{
    "model": "local-model",
    "messages": [{"role": "user", "content": "Hello"}]
  }'
```
