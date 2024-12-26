import fitz
from tqdm import tqdm
from bs4 import BeautifulSoup
import os
import json
import sys

def process_pdf(in_path :str):
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

    txt = process_pdf(in_path)

    # print(txt)

    file_name = os.path.basename(file_path).split('.')[-2]

    out_path = os.path.join(os.path.dirname(file_path), f"{file_name}.txt")

    print(f"正在写入到 {out_path}")

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(txt)

    print(f"结果已写入到 {out_path}")

    input('按Enter键关闭程序')



if __name__ == '__main__':
    main()