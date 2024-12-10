import os
import comtypes.client
import re
import traceback


def sanitize_filename(filename):
    """
    替换文件名中的特殊字符为下划线。
    """
    return re.sub(r'[\\/*?:"<>|（）()]', "_", filename)

def ppt_to_pdf(input_folder, output_folder):
    """
    将文件夹中的所有PPT文件转换为PDF文件。

    :param input_folder: 输入文件夹路径，包含PPT文件
    :param output_folder: 输出文件夹路径，用于存放PDF文件
    """
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 初始化 PowerPoint 应用
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1

    for filename in os.listdir(input_folder):
        if filename.endswith(".ppt") or filename.endswith(".pptx"):
            sanitized_name = sanitize_filename(filename)  # 替换特殊字符后的文件名
            ppt_path = os.path.abspath(os.path.join(input_folder, filename))
            pdf_path = os.path.abspath(os.path.join(output_folder, os.path.splitext(sanitized_name)[0] + ".pdf"))

            print(f"完整路径: {ppt_path} (长度: {len(ppt_path)})")

            
            try:
                print(f"正在处理文件: {ppt_path}")  # 调试信息
                presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
                presentation.SaveAs(pdf_path, 32)  # 32代表PDF格式
                presentation.Close()
                print(f"转换成功: {filename} -> {os.path.basename(pdf_path)}")
            except Exception as e:
                print(f"转换失败: {filename}, 错误: {e}")
                print(traceback.format_exc())

    # 退出 PowerPoint 应用
    powerpoint.Quit()

if __name__ == "__main__":
    input_folder = r"你的输入文件夹路径"  # 替换为你的输入文件夹路径
    output_folder = r"你的输出文件夹路径"  # 替换为你的输出文件夹路径

    ppt_to_pdf(input_folder, output_folder)
