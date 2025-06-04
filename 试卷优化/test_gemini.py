import google.generativeai as genai
import os
import time

# --- 配置 ---
# 替换为你的 Gemini API 密钥
# 强烈建议从环境变量读取，例如:
# GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
GEMINI_API_KEY = "AIzaSyCyyGPJoNr0HlcPvZQ8cg3ItbxFPv-Q_QY"

def run_simple_gemini_test():
    """
    执行一个简单的 Gemini API 调用测试。
    """
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 开始简单的 Gemini API 测试...")

    # 1. 配置 API
    try:
        if not GEMINI_API_KEY or GEMINI_API_KEY == "YOUR_GEMINI_API_KEY":
            print("错误：请在脚本中设置你的 GEMINI_API_KEY。")
            return
        genai.configure(api_key=GEMINI_API_KEY)
        # 你可以尝试 'gemini-1.0-pro' 或 'gemini-1.5-flash-latest'
        model_name = 'gemini-1.5-flash-latest'
        model = genai.GenerativeModel(model_name)
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Gemini API 配置成功，使用模型: {model.model_name}")
    except Exception as e:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 配置 Gemini API 时发生错误: {e}")
        return

    # 2. 准备一个简单的 Prompt
    simple_prompt = "你好，请用一句话告诉我今天星期几？" # 或者更简单的 "你好"
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 准备发送简单提示: '{simple_prompt}'")

    # 3. 调用 API
    try:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 正在调用 model.generate_content()...")
        # 为了更好地观察，我们可以增加一个简单的超时机制（通过外部线程，这里简化不加）
        # 或者设置 generation_config 中的参数（但 generate_content 本身是阻塞的）

        response = model.generate_content(simple_prompt) # 这是阻塞调用

        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] model.generate_content() 调用已返回。")

        # 4. 处理响应
        if not response.candidates:
            print("Gemini API 未返回有效候选内容。可能原因：")
            if hasattr(response, 'prompt_feedback') and response.prompt_feedback:
                for i, rating in enumerate(response.prompt_feedback.safety_ratings):
                     print(f"  - Prompt Safety Rating {i}: Category: {rating.category}, Probability: {rating.probability}")
            else:
                print("  - 没有详细的 prompt_feedback。")
            print(f"  - 查看 response.parts 是否有信息: {response.parts if hasattr(response, 'parts') else 'N/A'}")
            return

        response_text = ""
        if hasattr(response, 'text') and response.text:
            response_text = response.text
        elif response.parts:
            response_text = "".join(part.text for part in response.parts if hasattr(part, 'text'))

        if response_text.strip():
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Gemini API 简单测试成功！响应:")
            print(f"'{response_text.strip()}'")
        else:
            print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Gemini API 响应文本为空。")
            print(f"Full API Response object: {response}")


    except Exception as e:
        print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 调用 Gemini API 时发生错误: {e}")
        # 打印更详细的错误信息，如果可用
        if hasattr(e, 'args') and e.args:
            for arg_idx, arg_val in enumerate(e.args):
                print(f"  Error Arg [{arg_idx}]: {arg_val}")
                if hasattr(arg_val, 'message'): # google.api_core.exceptions.GoogleAPICallError
                    print(f"    Google API Error Message: {arg_val.message}")
        # 如果是 requests.exceptions.RequestException 或 urllib3 相关的错误，也可能在这里捕获到
        # 例如，网络连接问题
        if "Max retries exceeded with url" in str(e) or "Connection timed out" in str(e):
            print("  检测到可能的网络连接问题或超时。")


if __name__ == "__main__":
    run_simple_gemini_test()
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 简单测试脚本执行完毕。")