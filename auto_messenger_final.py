import uiautomation as auto
import time
import dashscope
from dashscope import MultiModalConversation
import os
import pyautogui
import schedule
import atexit
import signal
from PIL import Image, ImageDraw, ImageFont
import base64
from io import BytesIO
from dotenv import load_dotenv
from datetime import datetime

# ====================== 全局配置项 ======================
# 截图命名配置
TEST_SCREENSHOT_PREFIX = "test_small_area_"
SCREENSHOT_PREFIX = "wechat_screenshot_"
CHAT_SCREENSHOT_PREFIX = "chat_screenshot_"
AI_REPLY_SCREENSHOT_PREFIX = "ai_reply_screenshot_"
SCREENSHOT_SUFFIX = ".png"

SMALL_AREA_LEFT = 80
SMALL_AREA_TOP = 230
SMALL_AREA_RIGHT = 130
SMALL_AREA_BOTTOM = 290
RED_THRESHOLD = 10  # 红色检测阈值（越低越灵敏）

# 通义千问API配置
DASHSCOPE_API_KEY = ""

# 定时任务间隔（分钟）
INTERVAL_MINUTES = 1

# AI通用回复人设提示词（新增思考过程要求）
AI_PERSONA = """
你是清华大学计算机系大一的男生，性格开朗、友善、接地气，熟悉大学生日常聊天方式。
请严格按照以下两步执行：
第一步（思考过程）：详细说明你对聊天上下文的理解、回复的思路（比如语气、内容、贴合点）；
第二步（最终回复）：仅输出符合要求的回复内容，无任何多余解释。
思考过程和最终回复用【思考】和【回复】标签分隔，格式示例：
【思考】用户问我有没有空去打球，我需要用大学生的口语回复，自然一点，比如“刚写完作业～可以啊，几点？”
【回复】刚写完作业～可以啊，几点？

回复要求：
1. 深度理解截图中最新的对话上下文，回复要贴合语境，不答非所问；
2. 语气口语化、自然，像真人聊天（比如用“哈哈哈”“嗯嗯”“刚忙完～”等口语词）；
3. 避免机械感，不要使用书面化、官方化的表达，拒绝“好的”“收到”这类生硬回复；
4. 适配所有朋友关系，回复长度控制在1-100字，仅【回复】部分会被发送给好友。
"""

# ====================== 初始化与工具函数 ======================
def init_api_config():
    """初始化通义千问API配置"""
    load_dotenv()
    api_key = os.getenv("DASHSCOPE_API_KEY", DASHSCOPE_API_KEY)
    if not api_key:
        raise ValueError("请先配置通义千问API-KEY！")
    dashscope.api_key = api_key

def image_to_base64(image_path):
    """将本地图片转换为Base64编码"""
    try:
        with Image.open(image_path) as img:
            supported_formats = ["JPEG", "PNG", "WEBP", "BMP"]
            if img.format not in supported_formats:
                raise ValueError(f"不支持的图片格式：{img.format}，仅支持{supported_formats}")
            
            buffer = BytesIO()
            img.save(buffer, format=img.format)
            base64_str = base64.b64encode(buffer.getvalue()).decode("utf-8")
            return f"data:image/{img.format.lower()};base64,{base64_str}"
    except FileNotFoundError:
        raise FileNotFoundError(f"图片文件不存在：{image_path}")
    except Exception as e:
        raise Exception(f"图片转换Base64失败：{str(e)}")

def clean_all_screenshots():
    """清理所有生成的截图（主界面+聊天界面+AI回复截图）"""
    try:
        deleted_count = 0
        for file in os.listdir("."):
            if (file.startswith(SCREENSHOT_PREFIX) or 
                file.startswith(CHAT_SCREENSHOT_PREFIX) or
                file.startswith(AI_REPLY_SCREENSHOT_PREFIX)) and file.endswith(SCREENSHOT_SUFFIX):
                os.remove(file)
                print(f"[{datetime.now()}] 已清理截图文件：{file}")
                deleted_count += 1
        if deleted_count == 0:
            print(f"[{datetime.now()}] 无残留截图需要清理")
        else:
            print(f"[{datetime.now()}] 共清理 {deleted_count} 个截图文件")
    except Exception as e:
        print(f"[{datetime.now()}] 清理截图出错：{str(e)}")

def get_and_activate_wechat_window():
    """
    核心修复：兼容所有uiautomation版本的微信窗口激活逻辑
    移除IsMinimized，改用兼容方式恢复窗口
    """
    # 1. 查找微信窗口（兼容中文/英文名称）
    wechat_window = auto.WindowControl(searchDepth=1, Name="微信")
    if not wechat_window.Exists():
        wechat_window = auto.WindowControl(searchDepth=1, Name="WeChat")
    if not wechat_window.Exists():
        raise Exception("❌ 未找到微信窗口，请确保微信已打开并登录")
    
    # 2. 兼容版：直接恢复窗口（不管是否最小化，都执行ShowNormal）
    try:
        wechat_window.ShowNormal()  # 取消最小化/还原窗口
        time.sleep(0.8)
    except:
        pass  # 版本不兼容时忽略，不影响核心逻辑
    
    # 3. 确保窗口前置、激活
    wechat_window.SetTopmost(True)
    wechat_window.SetActive()
    time.sleep(1)
    wechat_window.SetTopmost(False)
    
    print(f"[{datetime.now()}] ✅ 微信窗口已恢复并前置显示")
    return wechat_window

def capture_wechat_screenshot():
    """精准截取微信主窗口（修复IsMinimized属性缺失问题）"""
    try:
        # 第一步：激活并恢复微信窗口（兼容版逻辑）
        wechat_window = get_and_activate_wechat_window()
        
        # 第二步：获取微信窗口坐标（兼容所有版本）
        rect = wechat_window.BoundingRectangle
        x = rect.left
        y = rect.top
        width = rect.right - rect.left
        height = rect.bottom - rect.top
        
        # 第三步：生成截图并保存
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = f"{SCREENSHOT_PREFIX}{timestamp}{SCREENSHOT_SUFFIX}"
        
        # 仅截取微信窗口区域
        screenshot = pyautogui.screenshot(region=(x, y, width, height))
        screenshot.save(screenshot_path)
        
        print(f"[{datetime.now()}] ✅ 微信主界面截图成功：{screenshot_path}")
        return screenshot_path
    
    except Exception as e:
        raise Exception(f"主界面截图失败：{str(e)}")

def capture_chat_screenshot():
    """截取当前打开的微信聊天界面"""
    try:
        wechat_window = get_and_activate_wechat_window()
        rect = wechat_window.BoundingRectangle
        x = rect.left
        y = rect.top
        width = rect.right - rect.left
        height = rect.bottom - rect.top

        # 生成聊天截图文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        chat_screenshot_path = f"{CHAT_SCREENSHOT_PREFIX}{timestamp}{SCREENSHOT_SUFFIX}"
        
        # 截取聊天窗口区域
        screenshot = pyautogui.screenshot(region=(x, y, width, height))
        screenshot.save(chat_screenshot_path)
        
        print(f"[{datetime.now()}] ✅ 微信聊天界面截图成功：{chat_screenshot_path}")
        return chat_screenshot_path
    except Exception as e:
        raise Exception(f"聊天界面截图失败：{str(e)}")

def generate_text_screenshot(text, filename_prefix):
    """将文本生成图片（截取AI回复内容）"""
    try:
        # 设置图片尺寸和字体
        font_size = 20
        # 尝试加载系统字体，兼容不同环境
        try:
            font = ImageFont.truetype("simhei.ttf", font_size)  # 黑体
        except:
            font = ImageFont.load_default(size=font_size)
        
        # 计算文本尺寸
        text_lines = text.split('\n')
        max_width = max([font.getbbox(line)[2] for line in text_lines])
        total_height = sum([font.getbbox(line)[3] for line in text_lines]) + 20  # 加边距
        
        # 创建图片并绘制文本
        img = Image.new("RGB", (max_width + 20, total_height), color="white")
        draw = ImageDraw.Draw(img)
        y_offset = 10
        for line in text_lines:
            draw.text((10, y_offset), line, font=font, fill="black")
            y_offset += font.getbbox(line)[3]
        
        # 保存图片
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = f"{filename_prefix}{timestamp}{SCREENSHOT_SUFFIX}"
        img.save(screenshot_path)
        
        print(f"[{datetime.now()}] ✅ AI回复文本截图成功：{screenshot_path}")
        return screenshot_path
    except Exception as e:
        raise Exception(f"生成文本截图失败：{str(e)}")
    
# ====================== 红色像素检测（未读消息识别） ======================
def capture_small_area_and_check_red():
    """截取指定小区域并检测红色像素（判断是否有未读消息）"""
    print("="*60)
    print(f"[{datetime.now()}] 开始检测未读消息红点")
    print("="*60)
    
    try:
        # 激活微信窗口并获取坐标
        wechat_window = get_and_activate_wechat_window()
        win_rect = wechat_window.BoundingRectangle
        win_x, win_y = win_rect.left, win_rect.top
        win_width, win_height = win_rect.right - win_rect.left, win_rect.bottom - win_rect.top
        
        # 计算检测区域绝对坐标（增加边界检查）
        small_x = win_x + SMALL_AREA_LEFT
        small_y = win_y + SMALL_AREA_TOP
        small_width = SMALL_AREA_RIGHT - SMALL_AREA_LEFT
        small_height = SMALL_AREA_BOTTOM - SMALL_AREA_TOP
        
        # 边界检查
        if small_x + small_width > win_x + win_width:
            small_width = win_x + win_width - small_x
        if small_y + small_height > win_y + win_height:
            small_height = win_y + win_height - small_y
        if small_width <= 0 or small_height <= 0:
            raise Exception(f"检测区域超出微信窗口范围！窗口尺寸：{win_width}x{win_height}")
        
        print(f"📌 检测区域：")
        print(f"   微信窗口相对：({SMALL_AREA_LEFT},{SMALL_AREA_TOP}) → ({SMALL_AREA_RIGHT},{SMALL_AREA_BOTTOM})")
        print(f"   屏幕绝对坐标：({small_x},{small_y}) 尺寸：{small_width}×{small_height}px")
        
        # 截取小区域并保存
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = f"{TEST_SCREENSHOT_PREFIX}{timestamp}{SCREENSHOT_SUFFIX}"
        time.sleep(0.5)  # 等待窗口稳定
        small_screenshot = pyautogui.screenshot(region=(small_x, small_y, small_width, small_height))
        small_screenshot.save(screenshot_path)
        print(f"✅ 检测区域截图已保存：{os.path.abspath(screenshot_path)}")
        
        # 检测红色像素
        red_pixel_count = 0
        img_rgb = small_screenshot.convert("RGB")
        pixels = img_rgb.load()
        for x in range(small_width):
            for y in range(small_height):
                r, g, b = pixels[x, y]
                # 红色判定：R值远大于G/B
                if r > RED_THRESHOLD and r > g * 1.5 and r > b * 1.5:
                    red_pixel_count += 1
        
        # 输出检测结果
        if red_pixel_count > 0:
            print(f"🚨 检测到 {red_pixel_count} 个红色像素（有未读消息）")
            has_red = True
        else:
            print(f"✨ 未检测到红色像素（无未读消息）")
            has_red = False
        
        return has_red
    
    except Exception as e:
        print(f"❌ 红点检测失败：{str(e)}")
        return False

# ====================== AI 回复生成（新增思考过程+文本截图） ======================
def generate_chat_reply(image_path):
    """根据聊天界面截图，生成通用、自然的回复（含思考过程+Token统计+文本截图）"""
    # 初始化API配置
    init_api_config()
    
    # 将截图转为Base64编码
    img_base64 = image_to_base64(image_path)
    
    # 构造AI请求消息
    messages = [{
        "role": "user",
        "content": [
            {"type": "text", "text": AI_PERSONA},
            {"type": "image", "image": img_base64}
        ]
    }]
    
    # 调用通义千问多模态API
    response = MultiModalConversation.call(
        model="qwen-vl-plus",
        messages=messages,
        result_format="text",
        temperature=0.7,  # 回复多样性（0-1，值越高越灵活）
        top_p=0.95,
        max_tokens=256     # 增大token上限，容纳思考过程
    )
    
    # 校验API响应状态
    if response.status_code != 200:
        raise Exception(f"AI回复生成失败：{response.code} - {response.message}")
    
    # 解析Token消耗（关键：新增Token统计）
    usage = response.usage
    prompt_tokens = usage.get('input_tokens', 0)  # 输入消耗Token
    completion_tokens = usage.get('output_tokens', 0)  # 输出消耗Token
    total_tokens = usage.get('total_tokens', 0)  # 总消耗Token
    print(f"[{datetime.now()}] 💰 Token消耗：输入{prompt_tokens} | 输出{completion_tokens} | 总计{total_tokens}")
    
    # 解析AI回复内容
    content = response.output.choices[0].message.content
    if isinstance(content, list):
        full_response = ''.join([item.get('text', '') for item in content if 'text' in item]).strip()
    else:
        full_response = str(content).strip()
    
    # 提取思考过程和最终回复
    thinking_process = ""
    final_reply = ""
    if "【思考】" in full_response and "【回复】" in full_response:
        # 分割思考和回复
        thinking_part = full_response.split("【思考】")[1].split("【回复】")[0].strip()
        reply_part = full_response.split("【回复】")[1].strip()
        thinking_process = thinking_part
        final_reply = reply_part
    else:
        # 兼容未按格式回复的情况
        thinking_process = "未生成思考过程，直接回复"
        final_reply = full_response.strip()
    
    # 打印思考过程（仅日志，不发送）
    print(f"[{datetime.now()}] 🤔 AI思考过程：{thinking_process}")
    
    # 过滤空回复
    if not final_reply:
        raise Exception("AI未生成有效回复内容")
    
    # 生成AI最终回复的文本截图
    generate_text_screenshot(final_reply, AI_REPLY_SCREENSHOT_PREFIX)
    
    return final_reply

# ====================== 微信操作（修复重复发送问题） ======================
def refresh_wechat_search():
    """单独的刷新聊天框函数（抽离出来，避免重复调用）"""
    wechat_window = get_and_activate_wechat_window()
    print(f"[{datetime.now()}] 🔄 刷新微信搜索框")
    wechat_window.SendKeys('{Ctrl}f', waitTime=1)          # 打开微信搜索框
    wechat_window.SendKeys('', waitTime=1.5)         # 输入刷新使用账号昵称
    wechat_window.SendKeys('{Enter}', waitTime=1)          # 回车打开聊天窗口

def operate_wechat_send_message(contact_name, message=None):
    """
    自动操作微信发送消息（修复重复发送自动回复问题）
    :param contact_name: 好友昵称
    :param message: 要发送的内容，None则仅打开聊天窗口不发送
    """
    wechat_window = get_and_activate_wechat_window()
    
    # 仅在首次定位好友前刷新一次（避免重复）
    refresh_wechat_search()
    
    # 搜索并定位目标好友
    print(f"[{datetime.now()}] 🎯 正在定位好友：{contact_name}")
    wechat_window.SendKeys('{Ctrl}f', waitTime=1)          # 打开微信搜索框
    wechat_window.SendKeys(contact_name, waitTime=1.5)    # 输入好友昵称
    wechat_window.SendKeys('{Enter}', waitTime=1)          # 回车打开聊天窗口
    
    # 等待聊天界面加载完成
    time.sleep(1)
    
    # 仅在有消息要发送时，才输出自动回复前缀（核心修复：避免重复发送）
    if message is not None:
        # 发送自动回复前缀（仅发送一次）
        wechat_window.SendKeys('以下是自动回复:', waitTime=0.8)      
        wechat_window.SendKeys('{Enter}')                  
        time.sleep(1)
        
        # 发送AI生成的回复
        wechat_window.SendKeys(message, waitTime=0.8)      
        wechat_window.SendKeys('{Enter}')                  
        print(f"[{datetime.now()}] 📤 已向「{contact_name}」发送回复：{message}")

# ====================== 核心业务逻辑 ======================
def extract_wechat_unread_friend(image_path):
    """识别微信截图中第一个未读消息的好友（排除群聊）+ Token统计"""
    init_api_config()
    img_base64 = image_to_base64(image_path)
    
    prompt = """
    你需要分析这张微信界面截图，严格按照以下要求执行：
    1. 找到截图中最后一个带红色数字未读角标的聊天对象；
    2. 仅识别个人好友（排除所有群聊）；
    3. 只返回好友昵称，无其他内容；
    4. 无符合条件的好友返回"未检测到未读好友"。
    5. 不要返回任何不带红色未读数字未读角标的好友昵称！！！！！
    6. 如果你没有识别到未读好友，则"未检测到未读好友"；否则返回你第一个识别到的好友昵称，其他以日志形式呈现。
    """
    
    response = MultiModalConversation.call(
        model="qwen-vl-plus",
        messages=[{"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image", "image": img_base64}]}],
        result_format="text",
        temperature=0.1,
        top_p=0.9,
        max_tokens=100
    )
    
    if response.status_code == 200:
        # 解析Token消耗（识别好友环节也统计）
        usage = response.usage
        prompt_tokens = usage.get('input_tokens', 0)
        completion_tokens = usage.get('output_tokens', 0)
        total_tokens = usage.get('total_tokens', 0)
        print(f"[{datetime.now()}] 💰 识别未读好友Token消耗：输入{prompt_tokens} | 输出{completion_tokens} | 总计{total_tokens}")
        
        # 解析好友名称
        content = response.output.choices[0].message.content
        if isinstance(content, list):
            text_parts = [item.get('text', '') for item in content if 'text' in item]
            result = ''.join(text_parts).strip()
        else:
            result = str(content).strip()
        
        result = result.replace("：", "").replace(":", "").replace("用户名：", "").replace("联系人：", "").strip()
        return None if result == "未检测到未读好友" else result
    else:
        raise Exception(f"API调用失败：{response.code} - {response.message}")

def main_scheduled_task():
    """定时任务主逻辑：识别未读好友 → 截图聊天界面 → AI思考→生成回复 → 截图回复 → 自动发送"""
    try:
        print(f"\n{'='*60}")
        print(f"[{datetime.now()}] 开始执行定时任务（间隔{INTERVAL_MINUTES}分钟）")
        print(f"{'='*60}")
        
        # 步骤1：截取微信主界面，识别第一个未读个人好友
        main_screenshot_path = capture_wechat_screenshot()
        contact_name = extract_wechat_unread_friend(main_screenshot_path)
        
        if not contact_name:
            print(f"[{datetime.now()}] ❌ 未检测到有未读消息的个人好友，跳过发送")
            return
        
        print(f"[{datetime.now()}] ✅ 识别到第一个未读好友：{contact_name}")
        
        # 步骤2：打开该好友的聊天窗口（仅打开，不发送任何内容）
        operate_wechat_send_message(contact_name, message=None)
        
        # 步骤3：截取聊天界面
        chat_screenshot_path = capture_chat_screenshot()
        
        # 步骤4：调用AI生成回复（含思考过程+回复截图）
        reply_content = generate_chat_reply(chat_screenshot_path)
        print(f"[{datetime.now()}] ✅ AI最终回复：{reply_content}")
        
        # 步骤5：发送AI生成的回复（仅此时发送自动回复前缀+AI内容）
        operate_wechat_send_message(contact_name, message=reply_content)
        
    except Exception as e:
        print(f"[{datetime.now()}] ❌ 任务执行出错：{str(e)}")
    finally:
        # 清理所有截图（主界面+聊天界面+AI回复截图）
        clean_all_screenshots()

# ====================== 定时任务与退出处理 ======================
def start_scheduled_tasks():
    """启动定时任务"""
    # 注册退出清理钩子（正常退出/强制终止都清理截图）
    atexit.register(clean_all_screenshots)
    
    def handle_terminate_signal(signum, frame):
        print(f"\n[{datetime.now()}] 接收到终止信号，开始清理资源...")
        clean_all_screenshots()
        exit(0)
    
    # 处理Ctrl+C和系统终止信号
    signal.signal(signal.SIGINT, handle_terminate_signal)
    signal.signal(signal.SIGTERM, handle_terminate_signal)
    
    # 设置定时任务
    schedule.every(INTERVAL_MINUTES).minutes.do(main_scheduled_task)
    
    # 启动日志
    print(f"[{datetime.now()}] 🚀 微信自动化回复任务已启动")
    print(f"[{datetime.now()}] ⏰ 每隔{INTERVAL_MINUTES}分钟执行一次")
    print(f"[{datetime.now()}] ⚠️  按 Ctrl+C 可安全终止任务")

    # 持续循环：检测红色 → 有则执行任务 → 休眠 → 重复
    while True:
        # 1. 检测小区域是否有红色
        has_red = capture_small_area_and_check_red()
        
        if has_red:
            print(f"[{datetime.now()}] 🔴 检测到红色区域，执行核心任务...")
            # 2. 有红色则执行核心任务
            main_scheduled_task()
        else:
            print(f"[{datetime.now()}] 🟢 未检测到红色区域，跳过本次执行")
        
        # 3. 休眠指定分钟（添加进度条，不改动原有休眠时长）
        total_seconds = INTERVAL_MINUTES * 60
        progress_bar_length = 50  # 进度条长度（字符数）
        print(f"\n[{datetime.now()}] ⏳ 等待{INTERVAL_MINUTES}分钟后进行下一次检测...")
        
        for second in range(total_seconds):
            # 计算进度百分比和进度条填充长度
            progress = (second + 1) / total_seconds
            filled_length = int(progress_bar_length * progress)
            # 构建进度条字符串
            progress_bar = '█' * filled_length + '-' * (progress_bar_length - filled_length)
            # 计算剩余时间
            remaining_seconds = total_seconds - (second + 1)
            remaining_min = remaining_seconds // 60
            remaining_sec = remaining_seconds % 60
            # 打印进度条（覆盖当前行）
            print(f"\r[{datetime.now()}] [{progress_bar}] {progress*100:.1f}% 剩余时间：{remaining_min}分{remaining_sec}秒", end='', flush=True)
            # 休眠1秒
            time.sleep(1)
        
        # 进度条完成后换行
        print()

# ====================== 程序入口 ======================
if __name__ == "__main__":
    try:
        start_scheduled_tasks()
    except Exception as e:
        print(f"[{datetime.now()}] ❌ 程序异常终止：{str(e)}")
        clean_all_screenshots()
        print(f"[{datetime.now()}] ✅ 程序终止，所有截图已清理完成")