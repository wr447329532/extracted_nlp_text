from docx import Document
import os
import win32com.client
import dashscope
from http import HTTPStatus
import json

# 初始化一个空集合来存储已遇到的值
global_unique_keys = set()
# 全局字典用于存储去重后的键值对
global_results = {}
# 需要对doc文件转换为docx文件 该方法转换为docx格式后续有些问题 建议wps进行转换
def doc_to_docx(doc_path):
    # 确保文件路径存在
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"The file {doc_path} does not exist.")

    # 初始化Word应用程序
    word = win32com.client.Dispatch("Word.Application")
    word.visible = False  # 不显示Word应用程序窗口

    try:
        # 打开DOC文件
        doc = word.Documents.Open(doc_path)
        # 构造DOCX文件的路径
        docx_path = doc_path.rsplit('.', 1)[0] + '.docx'
        # 将文件另存为DOCX格式
        doc.SaveAs2(docx_path, FileFormat=16)  # FileFormat=16 表示DOCX格式
        doc.Close()
        return docx_path
    except Exception as e:
        print(f"Error converting {docx_path} to DOCX: {e}")
    finally:
        word.Quit()
# 对docx进行txt转换并保存
def docx_to_text(docx_path, text_file_path):
    # 加载docx文件
    doc = Document(docx_path)
    # 读取每个段落并将其内容添加到文本字符串中
    text = '\n'.join(paragraph.text for paragraph in doc.paragraphs)
    # 将文本保存到新文件
    with open(text_file_path, 'w', encoding='utf-8') as text_file:
        text_file.write(text)

#对文本进行分块
def process_file(text_file_path,chunk_size=1024):
    with open(text_file_path, 'r', encoding='utf-8') as file:
        while True:
            chunk = file.read(chunk_size)
            if not chunk:
                break
            text_extracted_keys(chunk)
            # 去重的函数


def text_extracted_keys(text_chunks):
    # 定义一个空的字典用来保存已经提取到的键值对
    response = dashscope.Understanding.call(
        model=dashscope.Understanding.Models.opennlu_v1,
        sentence=text_chunks,
        labels="项目名称、项目编号、招标人招标代理人、购买标书截止时间、投标截止时间、开标地址、项目预算、最高限价、质疑截止时间、质疑内容、保证金递交时间、保证金方式",
        api_key="sk-6b52fbaea55b4d3ca3168d9059766bcb",
        )

    if response.status_code == HTTPStatus.OK:
        # 处理输出
        text_content = response.output.get('text','')
        # 解析text_content中的内容为键值对
        pairs = [pair.split(": ") for pair in text_content.split(";") if pair]
        for pair in pairs:
            if len(pair) == 2 and pair[1] != "None":
                key, value = pair
                # 仅当键值对的值非空时保存或更新
                if key not in global_results:
                    global_results[key] = value
        extracted_data = {pair[0]: (None if pair[1] == "None" else pair[1]) for pair in pairs if len(pair) == 2}
        #print(extracted_data)

        # 打印去重后的结果，这里直接打印global_results而不是response.output
        #print(json.dumps(response.output, ensure_ascii=False, indent=4))
    else:
        print("Error:", response.code, response.message)
    print(json.dumps(global_results,ensure_ascii=False,indent=4))



if __name__ == '__main__':
    doc_path = "your file format .doc path"
    docx_path = "your file format .docx path"
    text_file_path = "your file format .txt path"
    #doc_to_docx(doc_path)
    docx_to_text(docx_path,text_file_path)
    process_file(text_file_path)