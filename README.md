# Azure AI Speech – PowerShell TTS & Avatar Batch Scripts

Batch-convert SSML files to **WAV audio** or **avatar video** using the [Azure AI Speech](https://learn.microsoft.com/azure/ai-services/speech-service/) REST API.  

Sample use cases:
 - Create audio narrations for PowerPoint slides by writing SSML files with the desired text, voice, and prosody settings.
 - Generate **talking avatar videos** — a photorealistic human speaking your text — for presentations, training materials, or advertisements.
 - Useful when you want to generate audio or video for multiple slides or content pieces at once.

Two authentication methods are provided for each feature — choose the one that fits your environment.

## Repository Structure

```
├── text-to-speech-apikey.ps1   # TTS script – API Key authentication
├── text-to-speech-entraid.ps1  # TTS script – Entra ID authentication
├── avatar-apikey.ps1           # Avatar script – API Key authentication
├── avatar-entraid.ps1          # Avatar script – Entra ID authentication
├── ssml/                       # Input  – place your SSML XML files here
│   └── sample.xml              # Example SSML file
├── audio/                      # Output – generated WAV files (auto-created)
└── video/                      # Output – generated avatar MP4 videos (auto-created)
```

## Prerequisites

### 1. Azure AI Speech Resource

1. Open the [Azure Portal](https://portal.azure.com/) and create an **Azure AI Speech** resource  
   (**Create a resource → AI + Machine Learning → Speech**).
2. Choose a **Region** (e.g. `swedencentral`, `eastus`) and a **Pricing tier**.
3. Note the following values from the resource's **Keys and Endpoint** page:

   | Value | Where to find it | Used by |
   |---|---|---|
   | **Region** | Overview / Keys and Endpoint | API Key script |
   | **Key 1** or **Key 2** | Keys and Endpoint | API Key script |
   | **Resource name** | Resource name at the top of the Overview page | Entra ID script |

4. **If you plan to use Entra ID authentication**, you must also enable a **Custom domain** on the resource:
   - Go to **Networking → Custom domain name** and set a unique subdomain.
   - Once enabled, the endpoint becomes `https://<resource-name>.cognitiveservices.azure.com`.

### 2. SSML Input Files

Place one or more `.xml` files in the `ssml/` folder. Each file must contain valid [SSML](https://learn.microsoft.com/azure/ai-services/speech-service/speech-synthesis-markup-structure).

Example (`ssml/sample.xml`):

```xml
<speak version="1.0" xml:lang="en-US">
  <voice name="en-US-JennyNeural">
    <prosody rate="0%" pitch="0%">
      Hello! This is a sample SSML file for Azure AI Speech text-to-speech.
    </prosody>
  </voice>
</speak>
```

> A full list of available voices can be found in the [Voice Gallery](https://speech.microsoft.com/portal/voicegallery).

---

## Text-to-Speech Scripts

### Option 1 – API Key Authentication

This is the simplest method. It uses the API key from the Speech resource directly.

### Setup

No additional modules are required.

### Configuration

Open `text-to-speech-apikey.ps1` and set the two variables at the top:

```powershell
$region = "<YOUR-REGION>"       # e.g. "swedencentral"
$apiKey = "<YOUR-API-KEY>"      # Key 1 or Key 2 from the Azure Portal
```

### Run

```powershell
.\text-to-speech-apikey.ps1
```

A progress bar shows the current file and overall progress.  
If a WAV file already exists, you will be prompted to choose:

| Key | Action |
|-----|--------|
| **S** | Skip this file |
| **O** | Overwrite this file |
| **A** | Skip All remaining duplicates |
| **L** | Overwrite All remaining duplicates |

> **Note:** The API key method uses the regional endpoint  
> `https://<region>.tts.speech.microsoft.com/cognitiveservices/v1`

---

### Option 2 – Microsoft Entra ID Authentication

Use this method when API key authentication is disabled on the resource (e.g. enforced by organizational policy).  
It authenticates with your Azure AD / Entra ID identity via a Bearer token.

### Setup

1. **Install the Az.Accounts PowerShell module** (one-time):

   ```powershell
   Install-Module Az.Accounts -Force -Scope CurrentUser -AllowClobber
   ```

2. **Sign in to Azure:**

   ```powershell
   Connect-AzAccount
   ```

   If the Speech resource is in a specific subscription, select it:

   ```powershell
   Set-AzContext -Subscription "<SUBSCRIPTION-ID>"
   ```

3. **Assign the RBAC role** on the Speech resource (requires an Owner or User Access Administrator):

   ```powershell
   New-AzRoleAssignment `
     -SignInName "<YOUR-EMAIL>" `
     -RoleDefinitionName "Cognitive Services Speech User" `
     -Scope "/subscriptions/<SUB-ID>/resourceGroups/<RG-NAME>/providers/Microsoft.CognitiveServices/accounts/<RESOURCE-NAME>"
   ```

   Alternatively, assign the role in the Azure Portal:
   - Navigate to the Speech resource → **Access control (IAM)** → **Add role assignment**.
   - Role: **Cognitive Services Speech User**
   - Assign to your user account.

4. **Enable a Custom domain** on the Speech resource (see Prerequisites above).  
   Entra ID authentication **only works with the custom-domain endpoint** — NOT the regional endpoint.

### Configuration

Open `text-to-speech-entraid.ps1` and set the variable at the top:

```powershell
$resourceName = "<YOUR-SPEECH-RESOURCE-NAME>"   # e.g. "my-speech-resource"
```

### Run

```powershell
.\text-to-speech-entraid.ps1
```

A progress bar shows the current file and overall progress.  
If a WAV file already exists, you will be prompted to choose:

| Key | Action |
|-----|--------|
| **S** | Skip this file |
| **O** | Overwrite this file |
| **A** | Skip All remaining duplicates |
| **L** | Overwrite All remaining duplicates |

> **Note:** The Entra ID method uses the custom-domain endpoint  
> `https://<resource-name>.cognitiveservices.azure.com/tts/cognitiveservices/v1`

---

## Avatar Video Scripts

The avatar scripts generate **talking avatar videos** — a photorealistic CG human speaking your SSML text.  
They use the [Azure AI Speech Avatar Batch Synthesis API](https://learn.microsoft.com/azure/ai-services/speech-service/batch-synthesis-avatar) (asynchronous: submit job → poll status → download video).

### Avatar Types

| Type | Description | `$avatarCustomized` | `$avatarStyle` |
|---|---|---|---|
| **Standard video avatar** | Pre-built characters with multiple styles (e.g. Lisa, Harry) at 1920×1080 | `$false` | Required |
| **Standard photo avatar** | Pre-built photo-based characters (e.g. Adrian, Amara) at 512×512 | `$false` | Not needed (leave empty) |
| **Custom photo avatar** | Your own avatar created from a real person's photo | `$true` | Not needed |

#### Standard Video Avatars (with styles)

| Character | Styles |
|---|---|
| Harry | business, casual, youthful |
| Jeff | business, formal |
| Lisa | casual-sitting, graceful-sitting, graceful-standing, technical-sitting, technical-standing |
| Lori | casual, graceful, formal |
| Max | business, casual, formal |
| Meg | formal, casual, business |

#### Standard Photo Avatars (no style needed)

Adrian, Amara, Amira, Anika, Bianca, Camila, Carlos, Clara, Darius, Diego, Elise, Farhan, Faris, Gabrielle, Hyejin, Imran, Isabella, Layla, Liwei, Ling, Marcus, Matteo, Rahul, Rana, Ren, Riya, Sakura, Simone, Zayd, Zoe

#### Custom Photo Avatar (from a real person's photo)

Creating a custom photo avatar requires a manual process with Microsoft:

1. **Apply for limited access** via the [intake form](https://aka.ms/customneural).
2. **Record a consent video** — the real person shown in the photo must provide verbal consent.
3. **Submit the photo and consent** to Microsoft, who will create the avatar model for you.
4. Once Microsoft has created the model, use its name as `$avatarCharacter` and set `$avatarCustomized = $true`.

> See [Custom text to speech avatar](https://learn.microsoft.com/azure/ai-services/speech-service/custom-avatar-create) for full details.

### Option 3 – Avatar with API Key Authentication

#### Configuration

Open `avatar-apikey.ps1` and set the variables at the top:

```powershell
$region          = "<YOUR-REGION>"      # e.g. "swedencentral"
$apiKey          = "<YOUR-API-KEY>"     # Key 1 or Key 2
$avatarCharacter = "lisa"               # Character name (see tables above)
$avatarStyle     = "graceful-sitting"   # Style (video avatars only)
$avatarCustomized = $false              # $true for custom photo avatar
$videoFormat     = "mp4"                # "mp4" or "webm"
$videoCodec      = "h264"              # "h264", "hevc", "vp9", or "av1"
$backgroundColor = "#00000000"          # RRGGBBAA (transparent by default)
$subtitleType    = "soft_embedded"      # "soft_embedded", "hard_embedded", "external_file", or "none"
```

#### Run

```powershell
.\avatar-apikey.ps1
```

### Option 4 – Avatar with Entra ID Authentication

Requires the same setup as the Entra ID TTS script above (Az.Accounts module, `Connect-AzAccount`, RBAC role, custom domain).

#### Configuration

Open `avatar-entraid.ps1` and set the variables at the top:

```powershell
$resourceName    = "<YOUR-SPEECH-RESOURCE-NAME>"
$avatarCharacter = "lisa"
$avatarStyle     = "graceful-sitting"
$avatarCustomized = $false
$videoFormat     = "mp4"
$videoCodec      = "h264"
$backgroundColor = "#00000000"
$subtitleType    = "soft_embedded"
```

#### Run

```powershell
.\avatar-entraid.ps1
```

### Avatar Output

Generated videos are saved in the `video/` folder:

```
ssml/slide1.xml  →  video/slide1.mp4
ssml/slide2.xml  →  video/slide2.mp4
```

Both avatar scripts display a progress bar and prompt on duplicate files, just like the TTS scripts.

> **Note:** Avatar batch synthesis is **asynchronous**. Each file submits a job and the script polls every 10 seconds until the video is ready. This is slower than TTS but produces video output.

---

## TTS Output

Generated WAV files are saved in the `audio/` folder with the same base name as the input SSML file:

```
ssml/slide1.xml  →  audio/slide1.wav
ssml/slide2.xml  →  audio/slide2.wav
```

The default output format is `riff-16khz-16bit-mono-pcm` (16 kHz, 16-bit, mono PCM WAV).  
You can change the `X-Microsoft-OutputFormat` header in the script to use a different format.  
See [supported audio formats](https://learn.microsoft.com/azure/ai-services/speech-service/rest-text-to-speech#audio-outputs) for all options.

## Authentication Comparison

| | API Key | Entra ID |
|---|---|---|
| **Ease of setup** | Simple — just paste the key | Requires module install, login, and RBAC |
| **Security** | Key can be leaked if committed to source control | No secrets in code — uses identity-based auth |
| **Endpoint (TTS)** | Regional (`<region>.tts.speech.microsoft.com`) | Custom domain (`<name>.cognitiveservices.azure.com`) |
| **Endpoint (Avatar)** | Regional (`<region>.api.cognitive.microsoft.com`) | Custom domain (`<name>.cognitiveservices.azure.com`) |
| **Custom domain required** | No | Yes |
| **RBAC role required** | No | Yes — "Cognitive Services Speech User" |
| **Works when API keys disabled** | No | Yes |

## License

MIT

## Author

[clement-lucas](https://github.com/clement-lucas)
