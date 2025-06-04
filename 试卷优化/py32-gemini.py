import os
import shutil
from pathlib import Path
import time

# --- 模块导入和可用性检查 ---
try:
    import win32com.client as win32
    from win32com.client import constants as word_constants
    PYWIN32_AVAILABLE = True
    print("pywin32 模块已加载。")
except ImportError:
    PYWIN32_AVAILABLE = False
    print("警告: pywin32 库未安装或无法导入，将无法处理 .doc 文件。")

try:
    from docx import Document as DocxDocument
    PYTHON_DOCX_AVAILABLE = True
    print("python-docx 模块已加载。")
except ImportError:
    PYTHON_DOCX_AVAILABLE = False
    print("警告: python-docx 库未安装或无法导入，将无法处理 .docx 文件。")

try:
    import PyPDF2
    PYPDF2_AVAILABLE = True
    print("PyPDF2 模块已加载。")
except ImportError:
    PYPDF2_AVAILABLE = False
    print("警告: PyPDF2 库未安装或无法导入，将无法处理 .pdf 文件。")

try:
    import google.generativeai as genai
    GOOGLE_GENERATIVEAI_AVAILABLE = True
    print("google-generativeai 模块已加载。")
except ImportError:
    GOOGLE_GENERATIVEAI_AVAILABLE = False
    print("警告: google-generativeai 库未安装或无法导入，无法使用 Gemini API。")


# --- 全局配置 ---
# !!! 重要: 请务必替换为你的真实 API Key !!!
GEMINI_API_KEY = "AIzaSyCyyGPJoNr0HlcPvZQ8cg3ItbxFPv-Q_QY"
# 或者从环境变量读取:
# GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")

# !!! 重要: 请修改为你的实际文件夹路径 !!!
# 假设你的脚本和这些文件夹在同一个父目录下，或者使用绝对路径
SCRIPT_DIR = Path(__file__).parent.resolve() # 获取脚本所在目录
INPUT_BASE_FOLDER = r"D:\学习\传热学\优学院资源试卷\东莞理工学院期末"
# INPUT_BASE_FOLDER = SCRIPT_DIR / "input_files"    # 输入文件夹，存放原始文档
OUTPUT_BASE_FOLDER = SCRIPT_DIR / "output_markdown" # 输出文件夹，存放转换后的MD文件

# 全局 Gemini 模型对象
gemini_model = None

# --- Gemini API 相关函数 ---
def configure_gemini_api(api_key):
    global gemini_model
    if not GOOGLE_GENERATIVEAI_AVAILABLE:
        print("错误: google-generativeai 库不可用。")
        return False
    try:
        if not api_key or api_key == "YOUR_GEMINI_API_KEY":
            print("错误：请在脚本中或通过环境变量设置你的 GEMINI_API_KEY。")
            return False
        genai.configure(api_key=api_key)
        model_name = 'gemini-1.5-flash-latest' # 速度快，上下文窗口大
        gemini_model = genai.GenerativeModel(model_name)
        print(f"Gemini API 配置成功，使用模型: {gemini_model.model_name}")
        return True
    except Exception as e:
        print(f"配置 Gemini API 时发生错误: {e}")
        gemini_model = None
        return False

def get_markdown_from_gemini(text_content, file_type_hint="文档", subject_knowledge="相关学科"):
    if not gemini_model:
        print("Gemini 模型未初始化，无法处理文本。")
        return None

    prompt = f"""
    请帮我处理一份从 {file_type_hint} 文件中提取的关于“{subject_knowledge}”学科的试卷、讲义或学习资料的原始文本。
    由于原始文件排版或提取过程可能引入问题，文本中可能包含错误、不必要的空格、乱码或结构混乱。
    需达成以下要求，最终输出为 Markdown 格式：

    1.  **精准识别与修正内容**：
        *   仔细阅读并理解原始文本，识别所有题目、章节、段落、列表、公式等内容元素。
        *   尽力修正因排版糟糕、字体重叠、OCR错误（如果适用）或提取问题导致的文本错误、乱码和不通顺之处。
        *   补全因格式问题可能丢失的上下文或结构信息，使其逻辑连贯。

    2.  **Markdown 结构**：
        *   **文档标题**：如果能从文本中识别出主标题，请使用 Markdown 一级标题 (`# 标题`)。
        *   **章节/题型标题**：主要的章节标题或大的题型分类（如“第一章绪论”、“一、选择题”）使用 Markdown 二级标题 (`## 标题`)。更细分的子章节或小题型可使用三级标题 (`### 标题`)。
        *   **段落**：普通文本段落自然转换，段落间用一个空行分隔。

    3.  **特定元素格式**：
        *   **列表**：原文中的有序列表（如 1., 2., a., b.）和无序列表（如项目符号）应转换为对应的 Markdown 列表。
        *   **填空题**（如果文档包含）：题号后跟题干，填空处使用 `____` (连续四个下划线) 表示。示例：`1. 热力学第二定律的克劳修斯表述是 ____。`
        *   **判断题**（如果文档包含）：题号后跟题干，末尾用 `( )` 预留作答空间。示例：`1. 可逆过程一定是准静态过程。 ( )`
        *   **选择题**（如果文档包含）：题号后跟题干，每个选项另起一行，以大写字母和点号开头（如 `A.`, `B.`）。
            ```
            1. 下列哪个过程是不可逆的？
            A. 理想气体的自由膨胀
            B. 卡诺循环
            C. 缓慢的等温压缩
            ```
        *   **数学公式**：所有数学公式必须使用 Markdown 支持的 LaTeX 语法。
            *   行内公式使用单个美元符号包裹，例如：`$E = mc^2$`。
            *   独立的块级公式使用双美元符号包裹，例如：
                `$$ Q = \int T dS $$`
            *   确保所有符号（希腊字母、上下标、积分号、求和号等）都正确转换为 LaTeX 格式。
        *   **代码块**（如果文档包含）：程序代码或命令行示例应使用 Markdown 的代码块（三个反引号 ```）包裹，并尽可能指明语言类型。
        *   **表格**（如果文档包含）：如果能识别出表格结构，请尝试转换为 Markdown 表格。如果表格过于复杂，可以用文字清晰描述表格内容和结构。

    4.  **图片/图示处理**：
        *   如果文本中明确提及或暗示了图片、图表或示意图，并且你能从上下文中理解其内容，请用详细的文字描述替代。描述应以 `[图示描述：...]` 的格式给出。

    5.  **内容分隔**：
        *   在每道独立的题目之间（例如，选择题1和选择题2之间）、每个主要章节之间，或逻辑上独立的内容块之间，请确保其后有**一个清晰的空行**（即在上一块内容的最后一行之后，按两次Enter键），以形成自然的垂直分隔，提高可读性。

    6.  **输出要求**：
        *   生成的 Markdown 代码必须语法正确，结构清晰，能在常见的 Markdown 编辑器和渲染器中良好显示。
        *   专注于内容的准确性和完整性，尽可能保留所有重要信息，同时优化排版和表达。
        *   **请直接输出纯粹的 Markdown 代码内容，不要包含任何额外的开场白、解释、总结或对话性的文字。**

    以下是从 {file_type_hint} 文件中提取的关于“{subject_knowledge}”学科的原始文本内容：
    ---原始文本开始---
    {text_content}
    ---原始文本结束---

    请严格按照上述要求处理，并直接输出 Markdown 代码。
    """
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 正在向 Gemini API 发送请求 (文本长度: {len(text_content)} 字符, 主题: {subject_knowledge}, 文件类型提示: {file_type_hint})...")
    try:
        # 简单的调用方式，如果需要更细致的控制，可以添加 generation_config 和 safety_settings
        response = gemini_model.generate_content(prompt)

        if not response.candidates:
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Gemini API 未返回有效候选内容。可能原因：内容被安全设置阻止。")
            if hasattr(response, 'prompt_feedback') and response.prompt_feedback:
                for i, rating in enumerate(response.prompt_feedback.safety_ratings):
                     print(f"  - Prompt Safety Rating {i}: Category: {rating.category}, Probability: {rating.probability.name}") # 使用 .name 获取枚举名称
            return None

        markdown_result = ""
        if hasattr(response, 'text') and response.text:
            markdown_result = response.text
        elif response.parts: # 兼容一些可能没有 .text 但有 .parts 的情况
            markdown_result = "".join(part.text for part in response.parts if hasattr(part, 'text'))
        
        if not markdown_result.strip():
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Gemini API 响应的文本内容为空。")
            # print(f"Full API Response for debugging: {response}") # 取消注释以查看完整响应
            return None
            
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Gemini API 成功返回内容。")
        return markdown_result.strip()

    except Exception as e:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 调用 Gemini API 时发生严重错误: {e}")
        # 尝试打印更具体的错误信息
        if hasattr(e, 'args') and e.args:
            for arg_idx, arg_val in enumerate(e.args):
                print(f"  Error Arg [{arg_idx}]: {arg_val}")
                if hasattr(arg_val, 'message'):
                    print(f"    Google API Error Message: {arg_val.message}")
        return None

# --- 文本提取函数 ---
def extract_text_from_doc(doc_path: Path):
    if not PYWIN32_AVAILABLE:
        print(f"错误: pywin32 未安装，无法处理 .doc 文件: {doc_path.name}")
        return None
    word_app = None
    doc_obj = None
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 尝试使用 pywin32 打开 .doc 文件: {doc_path.name}")
    try:
        try:
            word_app = win32.gencache.EnsureDispatch('Word.Application')
        except AttributeError:
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] pywin32 Makepy 缓存可能存在问题。尝试不使用 gencache...")
            word_app = win32.Dispatch('Word.Application')
        
        word_app.Visible = False # 不显示 Word 应用
        doc_obj = word_app.Documents.Open(str(doc_path), ReadOnly=True)
        doc_obj.Activate()
        text_content = doc_obj.Content.Text
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 成功从 .doc 文件 '{doc_path.name}' 提取文本 (长度: {len(text_content)}).")
        return text_content
    except Exception as e:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 提取 .doc 文件 '{doc_path.name}' 内容时出错: {e}")
        if hasattr(e, 'com_error'): print(f"  COM Error: {e.com_error}") # type: ignore
        return None
    finally:
        if doc_obj:
            try: doc_obj.Close(word_constants.wdDoNotSaveChanges if hasattr(word_constants, 'wdDoNotSaveChanges') else 0)
            except Exception as e_close: print(f"  关闭文档 '{doc_path.name}' 时出错: {e_close}")
        if word_app:
            try: word_app.Quit()
            except Exception as e_quit: print(f"  退出Word应用时出错: {e_quit}")
        word_app = None # 确保 COM 对象被释放
        doc_obj = None

def extract_text_from_docx(docx_path: Path):
    if not PYTHON_DOCX_AVAILABLE:
        print(f"错误: python-docx 未安装，无法处理 .docx 文件: {docx_path.name}")
        return None
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 尝试使用 python-docx 打开 .docx 文件: {docx_path.name}")
    try:
        doc = DocxDocument(docx_path)
        full_text = [para.text for para in doc.paragraphs]
        text_content = '\n\n'.join(full_text) # 用双换行连接段落，更接近原始结构
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 成功从 .docx 文件 '{docx_path.name}' 提取文本 (长度: {len(text_content)}).")
        return text_content
    except Exception as e:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 提取 .docx 文件 '{docx_path.name}' 内容时出错: {e}")
        return None

def extract_text_from_pdf(pdf_path: Path):
    if not PYPDF2_AVAILABLE:
        print(f"错误: PyPDF2 未安装，无法处理 .pdf 文件: {pdf_path.name}")
        return None
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 尝试使用 PyPDF2 打开 .pdf 文件: {pdf_path.name}")
    text_content = ""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            if reader.is_encrypted:
                try:
                    reader.decrypt('') # 尝试用空密码解密
                    print(f"  PDF '{pdf_path.name}' 已用空密码解密。")
                except Exception as e_decrypt:
                    print(f"  警告: PDF '{pdf_path.name}' 已加密且无法用空密码解密: {e_decrypt}。提取的文本可能不完整或为空。")
                    # 对于有密码的PDF，这里可以添加密码输入逻辑，但目前跳过

            num_pages = len(reader.pages)
            print(f"  PDF '{pdf_path.name}' 包含 {num_pages} 页。")
            for page_num in range(num_pages):
                try:
                    page = reader.pages[page_num]
                    extracted_page_text = page.extract_text()
                    if extracted_page_text:
                        text_content += extracted_page_text + "\n\n" # 每页内容后加双换行
                    else:
                        print(f"  警告: PDF '{pdf_path.name}' 第 {page_num + 1} 页未提取到文本 (可能是图片或复杂布局)。")
                except Exception as e_page:
                    print(f"  提取 PDF '{pdf_path.name}' 第 {page_num + 1} 页时出错: {e_page}")
        
        text_content = text_content.strip()
        if not text_content:
            print(f"警告: 从 PDF '{pdf_path.name}' 提取的文本为空。该PDF可能主要是图片或扫描件，需要OCR处理。")
        else:
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 成功从 .pdf 文件 '{pdf_path.name}' 提取文本 (长度: {len(text_content)}).")
        return text_content
    except Exception as e:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 提取 .pdf 文件 '{pdf_path.name}' 内容时出错: {e}")
        return None

# --- 文件处理和保存 ---
def save_markdown_content(content: str, output_filepath: Path):
    try:
        output_filepath.parent.mkdir(parents=True, exist_ok=True)
        with open(output_filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Markdown 内容已成功保存到: {output_filepath}")
    except Exception as e:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 保存 Markdown 文件 '{output_filepath.name}' 时出错: {e}")

def process_single_file(file_path: Path, output_dir: Path, subject_knowledge="相关学科"):
    if not file_path.is_file():
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 跳过：'{file_path.name}' 不是一个文件。")
        return

    filename_stem = file_path.stem
    extension = file_path.suffix.lower()
    output_md_path = output_dir / f"{filename_stem}.md"

    # 如果输出文件已存在，可以选择跳过或覆盖。这里简单覆盖。
    # if output_md_path.exists():
    #     print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 跳过：输出文件 '{output_md_path.name}' 已存在。")
    #     return

    print(f"\n--- [{time.strftime('%Y-%m-%d %H:%M:%S')}] 开始处理文件: {file_path.name} ---")
    raw_text = None
    file_type_for_prompt = "文档"

    if extension == '.doc':
        file_type_for_prompt = "Word (.doc)"
        raw_text = extract_text_from_doc(file_path)
    elif extension == '.docx':
        file_type_for_prompt = "Word (.docx)"
        raw_text = extract_text_from_docx(file_path)
    elif extension == '.pdf':
        file_type_for_prompt = "PDF"
        raw_text = extract_text_from_pdf(file_path)
    else:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 跳过：不支持的文件类型 '{extension}' 对于文件 '{file_path.name}'。")
        return

    if raw_text is not None and raw_text.strip():
        if not gemini_model:
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 错误: Gemini API 未配置，无法处理 '{file_path.name}'。")
            return
        
        markdown_output = get_markdown_from_gemini(raw_text, file_type_hint=file_type_for_prompt, subject_knowledge=subject_knowledge)
        if markdown_output:
            save_markdown_content(markdown_output, output_md_path)
        else:
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 未能为 '{file_path.name}' 生成 Markdown 内容 (Gemini API 未返回有效结果)。")
    elif raw_text is None:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 未能从 '{file_path.name}' 提取文本，跳过 Gemini 处理。")
    else: # raw_text.strip() is empty
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 从 '{file_path.name}' 提取到的文本为空，跳过 Gemini 处理。")
    
    print(f"--- [{time.strftime('%Y-%m-%d %H:%M:%S')}] 文件 '{file_path.name}' 处理完毕 ---")


def batch_process_folder(input_folder_path_str: str, output_folder_path_str: str, subject_knowledge="相关学科", file_types_to_process=None):
    input_dir = Path(input_folder_path_str)
    output_dir = Path(output_folder_path_str)

    if not input_dir.is_dir():
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 错误：输入路径 '{input_dir}' 不是一个有效的文件夹。")
        return

    if not GOOGLE_GENERATIVEAI_AVAILABLE:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 错误: google-generativeai 库不可用，无法进行批量处理。")
        return

    if not configure_gemini_api(GEMINI_API_KEY):
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Gemini API 配置失败，批量处理中止。")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"\n=== [{time.strftime('%Y-%m-%d %H:%M:%S')}] 开始批量处理文件夹: {input_dir} ===")
    print(f"Markdown 输出将保存到: {output_dir}")

    if file_types_to_process is None: # 如果未指定，则处理所有支持的
        supported_extensions = []
        if PYWIN32_AVAILABLE: supported_extensions.append('.doc')
        if PYTHON_DOCX_AVAILABLE: supported_extensions.append('.docx')
        if PYPDF2_AVAILABLE: supported_extensions.append('.pdf')
        print(f"将处理以下所有支持的文件类型: {supported_extensions}")
    else:
        supported_extensions = [ft.lower().strip() for ft in file_types_to_process]
        print(f"将仅处理指定的文件类型: {supported_extensions}")

    if not supported_extensions:
        print("错误：没有可用的文件处理模块或未指定有效的文件类型。批量处理中止。")
        return

    processed_files_count = 0
    skipped_files_count = 0

    for item in input_dir.iterdir():
        if item.is_file():
            if item.suffix.lower() in supported_extensions:
                process_single_file(item, output_dir, subject_knowledge)
                processed_files_count +=1
            else:
                # print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 跳过 (文件类型不匹配或不支持): {item.name}")
                skipped_files_count +=1
        # 可以选择是否递归处理子文件夹，目前不处理
        # elif item.is_dir():
        #     print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 跳过子文件夹: {item.name}")

    print(f"\n=== [{time.strftime('%Y-%m-%d %H:%M:%S')}] 批量处理完成 ===")
    print(f"尝试处理文件数: {processed_files_count}")
    print(f"跳过（类型不符或非文件）数: {skipped_files_count}")

# --- 主执行流程 ---
if __name__ == "__main__":
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 脚本启动...")

    # 确保 GEMINI_API_KEY 已设置
    if not GEMINI_API_KEY or GEMINI_API_KEY == "YOUR_GEMINI_API_KEY":
        print("致命错误：请在脚本顶部设置 GEMINI_API_KEY。程序中止。")
        exit()

    # 再次确认并确保这些是 Path 对象
    # (即使之前的定义应该是正确的，这样做可以增加稳健性)
    current_script_dir = Path(__file__).parent.resolve()
    input_folder_path = Path(INPUT_BASE_FOLDER) # 如果 INPUT_BASE_FOLDER 已经是 Path 对象，Path() 调用是幂等的
    output_folder_path = Path(OUTPUT_BASE_FOLDER)

    # 如果 INPUT_BASE_FOLDER 和 OUTPUT_BASE_FOLDER 是相对于脚本目录定义的，
    # 确保它们是绝对路径或正确的相对路径
    if not input_folder_path.is_absolute():
        input_folder_path = current_script_dir / input_folder_path
    if not output_folder_path.is_absolute():
        output_folder_path = current_script_dir / output_folder_path


    print(f"输入文件夹将使用: {input_folder_path}")
    print(f"输出文件夹将使用: {output_folder_path}")

    # 确保基础输入输出文件夹存在
    try:
        input_folder_path.mkdir(parents=True, exist_ok=True)
        output_folder_path.mkdir(parents=True, exist_ok=True)
        print(f"确保文件夹存在: {input_folder_path} 和 {output_folder_path}")
    except Exception as e_mkdir:
        print(f"创建文件夹时发生错误: {e_mkdir}")
        print("请检查路径权限或路径是否有效。程序中止。")
        exit()


    if not any(input_folder_path.iterdir()): # 检查输入文件夹是否为空
        print(f"警告：输入文件夹 '{input_folder_path}' 为空。")
        print(f"请将需要转换的 .doc, .docx, .pdf 文件放入该文件夹中，然后重新运行脚本。")
    else:
        batch_process_folder(
            str(input_folder_path), # batch_process_folder 可能期望字符串路径
            str(output_folder_path),
            subject_knowledge="各类文档资料",
            file_types_to_process=['.doc', '.docx', '.pdf']
        )

    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 脚本执行完毕。")