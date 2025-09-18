import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from docx import Document
from openai import OpenAI
import requests
import msal

# ====== OpenRouter setup ======
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key="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"  # 👈 put your OpenRouter key here
)

# ====== Microsoft Graph setup ======
CLIENT_ID = "0a8aff2a-148b-4b77-9d18-af40f118ccce"   # from Azure App Registration
TENANT_ID = "c356ada5-b3f7-4a81-bee7-1657c780ad12"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["User.Read", "Files.ReadWrite", "offline_access"]

token_cache = msal.SerializableTokenCache()

def get_access_token():
    """Authenticate with Microsoft and return access token."""
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=token_cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    else:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise Exception("Device flow failed")
        messagebox.showinfo("Microsoft Login", flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Auth failed: {result}")

def list_onedrive_docs():
    """List Word docs in OneDrive root folder."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
    resp = requests.get(url, headers=headers).json()
    docs = [item for item in resp.get("value", []) if item["name"].endswith(".docx")]
    return docs

def download_file(file_id, save_as):
    """Download a OneDrive file by ID."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/content"
    resp = requests.get(url, headers=headers, stream=True)
    with open(save_as, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            f.write(chunk)

def upload_file(file_path, onedrive_folder="root"):
    """Upload summary back to OneDrive."""
    token = get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    filename = os.path.basename(file_path)
    url = f"https://graph.microsoft.com/v1.0/me/drive/{onedrive_folder}/children/{filename}/content"
    with open(file_path, "rb") as f:
        resp = requests.put(url, headers=headers, data=f)
    return resp.json()

# ====== Summarization logic ======
def read_word_documents(directory):
    """Read all .docx files from local folder and return combined text."""
    text_content = []
    for filename in os.listdir(directory):
        if filename.endswith(".docx"):
            filepath = os.path.join(directory, filename)
            doc = Document(filepath)
            for para in doc.paragraphs:
                if para.text.strip():
                    text_content.append(para.text)
    return "\n".join(text_content)

def summarise_text(text):
    """Send text to OpenRouter and get summary + corrections/lessons."""
    response = client.chat.completions.create(
        model="mistralai/mistral-7b-instruct:free",  # free model
        messages=[
            {
                "role": "system",
                "content": (
                    "You are a helpful assistant that summarizes documents into concise note form. "
                    "If there are errors, missing details, or lessons to be learned, identify them and fill in the gaps."
                )
            },
            {"role": "user", "content": text}
        ]
    )
    return response.choices[0].message.content.strip()

# ====== GUI Application ======
class SummarizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sumar (Test Build")
        self.root.geometry("750x550")

        # Buttons
        self.btn_select_local = tk.Button(root, text="Use Local Folder", command=self.select_local)
        self.btn_select_local.pack(pady=5)

        self.btn_select_onedrive = tk.Button(root, text="Use OneDrive Folder (Broken ATM)", command=self.select_onedrive)
        self.btn_select_onedrive.pack(pady=5)

        self.btn_select_save = tk.Button(root, text="Select Save Location", command=self.select_save)
        self.btn_select_save.pack(pady=5)

        self.btn_run = tk.Button(root, text="Run Summarizer", command=self.run_summarizer, state=tk.DISABLED)
        self.btn_run.pack(pady=5)

        # Output log
        self.output = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=20, width=90, state=tk.DISABLED)
        self.output.pack(pady=10)

        # Paths
        self.directory = None
        self.onedrive_files = None
        self.temp_downloads = []
        self.save_path = None

    def log(self, message):
        self.output.config(state=tk.NORMAL)
        self.output.insert(tk.END, message + "\n")
        self.output.see(tk.END)
        self.output.config(state=tk.DISABLED)
        self.root.update()

    def select_local(self):
        self.directory = filedialog.askdirectory()
        if self.directory:
            self.log(f"Selected local folder: {self.directory}")
        self.check_ready()

    def select_onedrive(self):
        try:
            self.onedrive_files = list_onedrive_docs()
            if not self.onedrive_files:
                messagebox.showinfo("OneDrive", "No Word documents found in OneDrive root folder.")
                return
            self.log("Downloading OneDrive docs...")
            os.makedirs("onedrive_temp", exist_ok=True)
            self.temp_downloads = []
            for file in self.onedrive_files:
                local_path = os.path.join("onedrive_temp", file["name"])
                download_file(file["id"], local_path)
                self.temp_downloads.append(local_path)
            self.log("OneDrive docs downloaded successfully.")
        except Exception as e:
            messagebox.showerror("OneDrive Error", str(e))
        self.check_ready()

    def select_save(self):
        self.save_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt")],
            title="Save summary as"
        )
        if self.save_path:
            self.log(f"Save path: {self.save_path}")
        self.check_ready()

    def check_ready(self):
        if (self.directory or self.temp_downloads) and self.save_path:
            self.btn_run.config(state=tk.NORMAL)

    def run_summarizer(self):
        if self.directory:
            self.log("Reading local Word documents...")
            text = read_word_documents(self.directory)
        elif self.temp_downloads:
            self.log("Reading downloaded OneDrive documents...")
            text_content = []
            for filepath in self.temp_downloads:
                doc = Document(filepath)
                for para in doc.paragraphs:
                    if para.text.strip():
                        text_content.append(para.text)
            text = "\n".join(text_content)
        else:
            messagebox.showerror("Error", "No documents selected.")
            return

        if not text.strip():
            messagebox.showerror("Error", "No text found in documents.")
            return

        self.log("Sending text to AI summarizer... please wait.")
        try:
            summary = summarise_text(text)
        except Exception as e:
            messagebox.showerror("Error", f"AI summarization failed:\n{e}")
            return

        self.log("Writing summary to file...")
        with open(self.save_path, "w", encoding="utf-8") as f:
            f.write(summary)

        self.log("✅ Done! Summary saved successfully.")
        messagebox.showinfo("Success", f"Summary saved at:\n{self.save_path}")

        # Offer to upload back to OneDrive
        if self.onedrive_files:
            if messagebox.askyesno("Upload", "Upload summary back to OneDrive?"):
                try:
                    result = upload_file(self.save_path)
                    self.log(f"Uploaded summary to OneDrive: {result.get('name', 'Unknown')}")
                except Exception as e:
                    messagebox.showerror("Upload Error", str(e))

def main():
    root = tk.Tk()
    app = SummarizerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

