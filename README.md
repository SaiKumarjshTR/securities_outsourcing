# SGML Pipeline API — Deployment Project

Thomson Reuters Securities SGML Pipeline packaged as a REST API for deployment on the TR AI Platform (Plexus).

---

## Project Structure

```
sgml-pipeline-deployment/
├── app/
│   ├── __init__.py
│   ├── config.py           # Environment-driven configuration
│   ├── main.py             # FastAPI application (routes)
│   ├── models.py           # Pydantic request/response models
│   └── pipeline_runner.py  # Shim that wraps batch_runner_standalone.py
├── pipeline/
│   └── batch_runner_standalone.py   ← COPY HERE (not in git)
├── data/
│   ├── COMPLETE_KEYING_RULES_UPDATED.txt   ← COPY HERE
│   ├── vendor_sgms/                         ← COPY HERE
│   └── chroma_db/                           ← generated at runtime
├── tests/
│   └── test_api.py
├── Dockerfile
├── requirements.txt
├── .env.example
└── .gitignore
```

---

## Prerequisites (one-time, on your Windows machine)

| # | Prerequisite | How |
|---|---|---|
| 1 | **WSL** | PowerShell (Admin): `wsl --install` → restart |
| 2 | **JFrog token** | VPN → [JFrog Self Service Portal](https://trten.sharepoint.com/sites/intr-artifactory-cop/SitePages/How-to-Log-in-to-Artifactory.aspx) → "Update or Rotate Token" |
| 3 | **MGMT credentials** | Your MGMT username + vault password |
| 4 | **DIA complete** | Required for Model Registry registration |

---

## Step 1: Prepare Files (Windows)

Copy the required files into the project before building:

```powershell
$SRC = "C:\Users\C303180\OneDrive - Thomson Reuters Incorporated\Desktop\TR\securities-outsourcing-samples\final_scripts"
$JURI = "C:\Users\C303180\OneDrive - Thomson Reuters Incorporated\Desktop\TR\securities-outsourcing-samples\sec-out-samples-2\Jurisdictions\juri"
$DST = "C:\Users\C303180\OneDrive - Thomson Reuters Incorporated\Desktop\TR\sgml-pipeline-deployment"

# 1. Pipeline script
Copy-Item "$SRC\batch_runner_standalone.py" "$DST\pipeline\"

# 2. Keying rules
Copy-Item "$SRC\COMPLETE_KEYING_RULES_UPDATED.txt" "$DST\data\"   # or from samples root
# (file is at ...\securities-outsourcing-samples\COMPLETE_KEYING_RULES_UPDATED.txt)

# 3. Create data dirs
New-Item -ItemType Directory -Force "$DST\data\vendor_sgms"
```

---

## Step 2: Git Repository

```bash
# In WSL terminal
cd /mnt/c/Users/C303180/OneDrive\ -\ Thomson\ Reuters\ Incorporated/Desktop/TR/sgml-pipeline-deployment
git init
git add .
git commit -m "Initial commit: SGML Pipeline API v0.0.1"

# Push to your TR GitHub / GitLab repository
git remote add origin https://github.com/tr-your-org/sgml-pipeline.git
git push -u origin main
```

---

## Step 3: Cloud-Tool Setup (WSL)

```bash
# In WSL Ubuntu terminal
python3 -m venv cloud-tool
source cloud-tool/bin/activate

pip3 install cloud-tool \
  --extra-index=https://{YOUR_EMPLOYEE_ID}:{YOUR_JFROG_TOKEN}@tr1.jfrog.io/tr1/api/pypi/pypi-local/simple
```

---

## Step 4: Docker Installation (WSL)

```bash
sudo apt update && sudo apt upgrade -y
sudo apt install -y apt-transport-https ca-certificates curl gnupg lsb-release
curl -fsSL https://download.docker.com/linux/ubuntu/gpg | sudo gpg --dearmor -o /usr/share/keyrings/docker-archive-keyring.gpg
echo "deb [arch=$(dpkg --print-architecture) signed-by=/usr/share/keyrings/docker-archive-keyring.gpg] https://download.docker.com/linux/ubuntu $(lsb_release -cs) stable" | sudo tee /etc/apt/sources.list.d/docker.list > /dev/null
sudo apt update && sudo apt install -y docker-ce docker-ce-cli containerd.io
sudo usermod -aG docker $USER
sudo service docker start
docker --version
```

---

## Step 5: AWS CLI + cloud-tool login (WSL)

```bash
curl "https://awscli.amazonaws.com/awscli-exe-linux-x86_64.zip" -o "awscliv2.zip"
sudo apt install -y unzip && unzip awscliv2.zip && sudo ./aws/install
aws --version
export AWS_PROFILE="default"
cloud-tool --region us-east-1 login   # provide MGMT username + vault password, select role 2/aiml
```

---

## Step 6: Build & Push Docker Image (WSL)

```bash
# From repo root in WSL
MODEL_NAME="sgml-pipeline-prod"
VERSION="0.0.1"
ECR="127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1"

# ECR login
aws ecr get-login-password --region us-east-1 | docker login --username AWS --password-stdin 127288631409.dkr.ecr.us-east-1.amazonaws.com

# Set JFrog credentials
export TR_JFROG_USERNAME="your_employee_id"
export TR_JFROG_TOKEN="your_jfrog_token"

# Build (with JFrog credentials passed as build args)
docker build --no-cache \
  -t ${MODEL_NAME} \
  --build-arg TR_JFROG_USERNAME="${TR_JFROG_USERNAME}" \
  --build-arg TR_JFROG_TOKEN="${TR_JFROG_TOKEN}" \
  --file Dockerfile .

# Test locally (optional)
docker run -p 8501:8501 ${MODEL_NAME}
# Visit http://localhost:8501/health

# Tag & push
docker tag ${MODEL_NAME} ${ECR}:${MODEL_NAME}-${VERSION}
docker push ${ECR}:${MODEL_NAME}-${VERSION}

# Copy this ARN for Model Registry:
echo "Image ARN: ${ECR}:${MODEL_NAME}-${VERSION}"
```

---

## Step 7: Model Registry Registration

URL: https://contentconsole.thomsonreuters.com/ai-platform/registry/model-registry/models

1. **Add Model**
   - Model Name: `sgml-pipeline-prod`
   - Git Repo Link: your git repository URL
   - DIA ID: your completed DIA ID

2. **Add Model Version**
   - Version: `0.0.1`
   - Tag: `0.0.1`
   - Development Platforms: `Desktop/Laptop`
   - Deployment Platform: `AI Platform`
   - Third Party Flag: `No`
   - Rule Based Flag: `No`
   - Machine Learning Task: `Text Generation`
   - Model Version Developers: your TR email
   - Image ARN: *(paste from Step 6)*

---

## Step 8: Deployment Activation

URL: https://contentconsole.thomsonreuters.com/ai-platform/deployment

1. Click **Create Deployment Job**
   - Model Name: `sgml-pipeline-prod`
   - Leave Environments & Secrets as default
2. Click **Save**
3. In Deployment Jobs → 3 dots (Actions) → **Activate Deployment Job**
4. Wait for status → **Active**
5. Click the deployment name → copy the **Deployment Endpoint URL** and **Health URL**

---

## API Usage

### Health Check
```
GET https://<deployment-id>.ai-platform.thomsonreuters.com/health
```

### Convert DOCX to SGML
```bash
curl -X POST "https://<deployment-id>.ai-platform.thomsonreuters.com/convert" \
  -H "accept: application/json" \
  -F "file=@your-document.docx"
```

Response:
```json
{
  "status": "success",
  "doc_name": "your-document",
  "sgml": "<?xml version=\"1.0\"?>...",
  "score": 97.4
}
```

---

## Model Name & Version

| Field | Value |
|---|---|
| Model Name | `sgml-pipeline-prod` |
| Version | `0.0.1` |
| Port | `8501` |
| Image ARN format | `127288631409.dkr.ecr.us-east-1.amazonaws.com/a207870-ml-model-registry-model-registry-prod-use1:sgml-pipeline-prod-0.0.1` |
