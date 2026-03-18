import tkinter as tk
from tkinter import filedialog, messagebox
import os
from dotenv import load_dotenv
from core.ppt_logic import make_report


# 1. BASE_DIR 먼저 정의 (프로젝트 루트)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 2. .env 경로 설정
env_path = os.path.join(BASE_DIR, ".env")

# 3. 존재 여부 확인
if not os.path.exists(env_path):
    print("⚠️ .env 파일을 찾을 수 없음:", env_path)

# 4. .env 로드
load_dotenv(env_path)

# 5. 환경변수 읽기
DEFAULT_TEMPLATE_PATH = os.getenv("DEFAULT_TEMPLATE_PATH", "")


def select_project_folder(entry):
    folder = filedialog.askdirectory()
    if folder:
        entry.delete(0, tk.END)
        entry.insert(0, folder)


def select_template_file(entry):
    file = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    if file:
        entry.delete(0, tk.END)
        entry.insert(0, file)


def run_generation(project_entry, template_entry):
    project_folder = project_entry.get()
    template_path = template_entry.get()

    if not os.path.exists(project_folder):
        messagebox.showerror("오류", "프로젝트 폴더를 선택하세요.")
        return

    if not template_path or not os.path.exists(template_path):
        messagebox.showerror("오류", "템플릿 파일 경로를 확인하세요.")
        return

    try:
        output = make_report(project_folder, template_path)
        messagebox.showinfo("완료", f"생성 완료: {os.path.basename(output)}")
    except Exception as e:
        messagebox.showerror("에러", str(e))


def run_ui():
    root = tk.Tk()
    root.title("PPT 자동 생성기")
    root.geometry("600x230")

    # 프로젝트 폴더
    tk.Label(root, text="프로젝트 폴더").pack(pady=5)
    frame1 = tk.Frame(root)
    frame1.pack()
    entry_project = tk.Entry(frame1, width=50)
    entry_project.pack(side=tk.LEFT)
    tk.Button(frame1, text="찾기",
              command=lambda: select_project_folder(entry_project)).pack(side=tk.LEFT)

    # 템플릿
    tk.Label(root, text="PPT 템플릿").pack(pady=5)
    frame2 = tk.Frame(root)
    frame2.pack()
    entry_template = tk.Entry(frame2, width=50)
    entry_template.pack(side=tk.LEFT)

    # 🔥 .env 값이 있으면 자동 입력
    if DEFAULT_TEMPLATE_PATH:
        entry_template.insert(0, DEFAULT_TEMPLATE_PATH)

    tk.Button(frame2, text="찾기",
              command=lambda: select_template_file(entry_template)).pack(side=tk.LEFT)

    # 실행 버튼
    tk.Button(root, text="보고서 생성", bg="blue", fg="white",
              command=lambda: run_generation(entry_project, entry_template)).pack(pady=20)

    root.mainloop()