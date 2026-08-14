import os
import sys

def install_and_import(package):
    try:
        __import__(package)
    except ImportError:
        print(f"'{package}' kitabxanasi tapilmadi. Qurasdirilir...")
        import subprocess
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        except Exception as e:
            print(f"Xeta: '{package}' kitabxanasini qurasdirmaq mumkun olmadi. Zehmet olmasa 'pip install {package}' emrini elle calisdirin. Xeta teferruati: {e}")
            sys.exit(1)

# Ensure huggingface_hub is installed
install_and_import("huggingface_hub")

from huggingface_hub import HfApi, login

TOKEN_FILE = "hf_token.txt"
REPO_ID = "Orxam91/Anbar_bot"

def get_token():
    env_token = os.environ.get("HF_TOKEN")
    if env_token and env_token.strip():
        return env_token.strip()

    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r", encoding="utf-8") as f:
            token = f.read().strip()
            if token:
                return token
                
    print("=" * 60)
    print("          HUGGING FACE YUKLEME KOMEKCISI")
    print("=" * 60)
    print("Fayllari yuklemek ucun Hugging Face 'Write' (Yazma) tokeni lazimdir.")
    print("Tokeni elde etmek ucun:")
    print("1. Bu linke kecin: https://huggingface.co/settings/tokens")
    print("2. 'New token' duymesine klikleyin.")
    print("3. Token tipini (Role) 'Write' secin ve ad verib (meselen: 'AnbarBot') yaradin.")
    print("4. Yaradilmis tokeni kopyalayib bura yapistirin (Sag klikleyib yapistira bilersiniz).")
    print("-" * 60)
    
    token = input("Hugging Face 'Write' Tokenini daxil edin: ").strip()
    if token:
        with open(TOKEN_FILE, "w", encoding="utf-8") as f:
            f.write(token)
    return token

def main():
    token = get_token()
    if not token:
        print("Xeta: Token daxil edilmedi. Yukleme legv olundu.")
        input("\nCixmaq ucun Enter duymesine basin...")
        return
        
    print(f"\nFayllar '{REPO_ID}' Space-ine yuklenir...")
    
    try:
        # Login
        login(token=token)
        
        # Initialize API, ensure repo exists and upload
        api = HfApi()
        try:
            api.create_repo(
                repo_id=REPO_ID,
                repo_type="space",
                space_sdk="docker",
                exist_ok=True,
                token=token
            )
        except Exception as repo_e:
            print(f"Space yaratmaq melumati: {repo_e}")

        api.upload_folder(
            folder_path="upload_to_huggingface",
            repo_id=REPO_ID,
            repo_type="space",
            token=token
        )
        print("\n" + "=" * 60)
        print("UGURLU! Butun fayllar Hugging Face Space-e yuklendi.")
        print("Tetbiq yeniden qurulur. Bir nece deqiqeye saytiniz aktiv olacaq.")
        print("=" * 60)
    except Exception as e:
        print(f"\nXeta bas verdi: {e}")
        print("\nQeyd: Eger tokeniniz yanlisdirsa ve ya sehv yazilibsa,")
        print("eyni qovluqdaki 'hf_token.txt' faylini silib proqrami yeniden basladin.")
        
    input("\nCixmaq ucun Enter duymesine basin...")

if __name__ == "__main__":
    main()
