from pprint import pprint
import fitz
from tqdm import tqdm
from bs4 import BeautifulSoup
import os
import json
import sys
import re
import xlsxwriter as xw


def write_to_xlsx(data: dict, fname: str):
    workbook = xw.Workbook(fname)  # 创建工作簿
    worksheet1 = workbook.add_worksheet("sheet1")  # 创建子表
    worksheet1.activate()  # 激活表
    title = ['页码', '中文标题', '英文标题', '作者_zh', '作者_en']  # 设置表头
    worksheet1.write_row('A1', title)  # 从A1单元格开始写入表头
    i = 2  # 从第二行开始写入数据
    for j in range(len(data)):
        insertData = [data[j]["页码"], data[j]["中文标题"], data[j]["英文标题"], data[j]["作者_zh"], data[j]["作者_en"]]
        row = 'A' + str(i)
        worksheet1.write_row(row, insertData)
        i += 1
    workbook.close()  # 关闭表
    pass


def is_chinese(char):
    if '\u4e00' <= char <= '\u9fff':
        return True
    else:
        return False

def process_pdf2(in_path: str) -> str:
    pdf = fitz.open(in_path) # 打开pdf文件
    pg_len = pdf.page_count
    txt = '页码\t中文标题\t英文标题\t作者_zh\t作者_en\n'
    dict_data = []
    count = 0

    zh_title_style = [(19, "FZSSK--GBK1-0"), (19, "E-BZ")] # 字体大小, 字形
    zh_title = ''


    zh_author_style = [(13, "FZKTK--GBK1-0")]
    zh_author_start = -1
    zh_author_end = 0


    en_title_style = [(13, "E-FZ"), (1, "E-BZ"), (13, "E-B6")]
    en_title = ''

    en_author_style = [(9, "E-BX")]
    en_author_start = -1
    en_author_end = 0
    
    for pg_num, page in tqdm(enumerate(pdf), total=pg_len):

        # 每个页面, 重新初始化 ====
        zh_title = ''
        zh_author_start = -1
        zh_author_end = 0
        en_title = ''
        en_author_start = -1
        en_author_end = 0

        # "文献标识码"是否出现
        literature_identification_code: bool = False

        # 初始化完毕 ====

        blks = page.get_text("dict")["blocks"]
        cnt = dict()

        all_data = []
        idx = 0
        for b in blks:
            # print('-'*100, b.keys())
            if 'lines' not in b.keys():
                continue
            for l in b["lines"]:
                for s in l["spans"]:
                    # print(s["size"], s["font"], s["text"])

                    all_data.append(s["text"])
                    idx += 1

                    if s["text"] == "文献标志码":
                        literature_identification_code = True

                    # 中文题目
                    if (int(s["size"]), s["font"]) in zh_title_style:
                        zh_title += s["text"]

                    # 中文作者
                    if (int(s["size"]), s["font"]) in zh_author_style:
                        # zh_author.append(s["text"])
                        if zh_author_start == -1:
                            zh_author_start = idx - 1
                        zh_author_end = idx - 1

                    # 英文题目
                    if (int(s["size"]), s["font"]) in en_title_style:
                        
                        # 如果文献标识码未出现, 一定不是英文题目
                        if literature_identification_code == False:
                            continue

                        if len(s["text"].strip()) == 0:
                            en_title += ' '
                        else:
                            en_title += s["text"]

                    # 英文作者
                    if (int(s["size"]), s["font"]) in en_author_style:
                        # 如果文献标识码未出现, 一定不是英文作者
                        if literature_identification_code == False:
                            continue
                        if en_author_start == -1:
                            en_author_start = idx - 1
                        en_author_end = idx - 1

                        # print('-' * 100, idx - 1)
                    

                    

        # pprint(cnt)

        # print(zh_title)
        # print(','.join(zh_author))
        # print(en_title)
                        
        zh_author = all_data[zh_author_start: zh_author_end + 1]
        zh_author_list = [i for i in zh_author if i == ',' or i == ' ' or is_chinese(i[0])]
        zh_author = re.sub(r',\s*,+', ',', ''.join(zh_author_list))

        # print(en_author_start, en_author_end)
        en_author_list = [i for i in all_data[en_author_start: en_author_end + 1] if i == ',' or i == ' ' or (not str(i[0]).isdigit() and len(i) > 1)]
        # print(en_author_list)
        tmp = []
        for i in range(len(en_author_list)):
            if i - 1 >= 0 and en_author_list[i] == ',' and en_author_list[i - 1] == ',':
                continue
            tmp.append(en_author_list[i])

        en_author_list = tmp
        en_author = ''.join(en_author_list)
        en_author = en_author.replace(", , ", ", ")
        # print(en_author)

        line_res = f"{pg_num + 1}\t{zh_title}\t{en_title}\t{zh_author}\t{en_author}\n"
        if literature_identification_code:
            txt += line_res
            dict_data.append({
                "页码": pg_num + 1,
                "中文标题": zh_title,
                "英文标题": en_title,
                "作者_zh": zh_author,
                "作者_en": en_author,
            })
            # print(zh_author)
            # print(zh_author)
    
    return txt, dict_data


def main2():
    in_path = "./src/测试用小半期论文PDF.pdf"
    txt = process_pdf2(in_path)
    pass


def process_pdf(in_path :str) -> str:
    """
    逐页面 将PDF转换为HTML字符串
    """
    pdf = fitz.open(in_path) # 打开pdf文件
    pg_len = pdf.page_count
    txt = '页码\t中文标题\t英文标题\t作者\n'
    count = 0
    # 遍历每一页
    for idx, page in tqdm(enumerate(pdf), total=pg_len):
        html = page.get_text('html')        # 转为html
        soup = BeautifulSoup(html, 'lxml')  # 使用bs4解析html

        # 获取中文title
        zh_titles = soup.find_all(style=lambda css: css and 'font-size:19.7pt' in css)
        # print(f"Page {idx + 1}")
        zh_title = ''.join([t.text for t in zh_titles])
        # print(f"题目: {zh_title}")


        # 获取中文作者
        zh_authors = soup.find_all(style=lambda css: css and 'font-family:FZKTK--GBK1,serif;font-size:13.1pt' in css)

        # 本页没有标题, 直接跳过
        if len(zh_authors) <= 0:
            continue

        zh_author = ','.join([a.text for a in zh_authors])
        # print(f"作者: {zh_author}")


        # 获取英文title
        en_titles = soup.find_all(style=lambda css: css and 'font-family:E,serif;font-size:13.1pt' in css)
        en_title = ' '.join([t.text for t in en_titles if t.text != ','])
        # print(f"题目en: {en_title}")

        txt += f"Page{idx + 1}\t{zh_title}\t{en_title}\t{zh_author}\n"
        count += 1
    
    print(f"共处理 {count} 篇论文")
    return txt

def main():
    args = sys.argv
    # print(args)

    if len(args) < 2:
        print(f"将要处理的PDF文件拖动至{args[0]}  或  使用命令: {args[0]} <待处理文件>")
        input('按任意键关闭程序')
        return

    file_path = args[1]

    if not os.path.exists(file_path):
        print(f"'{file_path}' 不存在, 程序已退出")
        input('按任意键关闭程序')
        return
    
    if not os.path.isfile(file_path) or not file_path.endswith(".pdf"):
        print(f"'{file_path}' 不是PDF文件, 程序已退出")
        input('按任意键关闭程序')
        return
    
    in_path = file_path

    if not os.path.isfile(in_path):
        print(f"{in_path} 文件不存在, 程序已退出")
        input('按任意键关闭程序')
        return

    print(f"处理 {file_path} 中......")

    txt, dict_data = process_pdf2(in_path)

    # print(txt)

    file_name = os.path.basename(file_path).split('.')[-2]

    out_path = os.path.join(os.path.dirname(file_path), f"{file_name}.txt")
    out_path2 = os.path.join(os.path.dirname(file_path), f"{file_name}.xlsx")


    print(f"\n正在写入到 {out_path}")
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(txt)
    print(f"结果已写入到 {out_path}\n")


    print(f"正在写入到 {out_path2}")
    write_to_xlsx(dict_data, out_path2)
    print(f"结果已写入到 {out_path2}")

    input('按Enter键关闭程序')



if __name__ == '__main__':
    main()