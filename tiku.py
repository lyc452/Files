import pandas as pd
import os
import re
from tqdm import tqdm
from datetime import datetime

# 列名映射字典（兼容不同列名）
COLUMN_MAPPING = {
    '题目': '题目内容',
    '问题': '题目内容',
    '题干': '题目内容',
    '正确选项': '答案',
    '正确答': '答案',
    '正确选择': '答案'
}

def clean_answer(answer, q_type):
    """答案规范化处理函数"""
    if pd.isna(answer):
        return ''
    
    # 统一转为大写并去除空白
    ans = str(answer).strip().upper()
    
    # 去除所有分隔符（中文顿号、逗号、空格等）
    ans = re.sub(r'[、,，\s]', '', ans)
    
    # 判断题处理
    if q_type == '判断题':
        patterns = {
            'A': [r'^T$', r'^正确$', r'^对$', r'^是$', r'√'],
            'B': [r'^F$', r'^错误$', r'^错$', r'^否$', r'×']
        }
        for key in patterns:
            for pattern in patterns[key]:
                if re.search(pattern, ans, re.IGNORECASE):
                    return key
    else:
        # 去除非选项字符（只保留A-F）
        ans = re.sub(r'[^A-F]', '', ans)
    
    return ans

def process_excel_files():
    # 输入输出路径
    source_folder = r"E:\工作\安全准入\2024年安全准入考试题库0218（专业列表修改版）\2024年安全准入考试题库0208（最终版）\安全准入考试题库0208"
    output_path = r"E:\工作\安全准入\准入考试20250414(全).xlsx"
    
    # 初始化结果集
    final_df = pd.DataFrame(columns=['题目内容', '答案', 'A', 'B', 'C', 'D', 'E', 'F'])
    error_log = []
    
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_file = os.path.join(script_dir, f"处理日志_{datetime.now().strftime('%Y%m%d%H%M')}.txt")

    # 遍历处理文件
    file_list = [f for f in os.listdir(source_folder) if f.lower().endswith(('.xls', '.xlsx'))]
    
    for file in tqdm(file_list, desc='处理进度', unit='file'):
        file_path = os.path.join(source_folder, file)
        
        try:
            # 读取Excel文件
            engine = 'xlrd' if file.lower().endswith('.xls') else 'openpyxl'
            df = pd.read_excel(
                file_path,
                header=0,
                engine=engine,
                keep_default_na=False,
                dtype=str
            )
            
            # 统一列名
            df.rename(columns=lambda x: COLUMN_MAPPING.get(x.strip(), x), inplace=True)
            
            # 数据清洗
            df['题目内容'] = df['题目内容'].astype(str).str.strip()
            df['题型'] = df['题型'].astype(str).str.strip()
            df['题型'] = df['题型'].apply(lambda x: '判断题' if '判断' in x else x)
            
            # 过滤无效数据
            df = df[df['题目内容'].str.len() > 3]
            
            # 处理每行数据
            for idx, row in df.iterrows():
                try:
                    q_type = row.get('题型', '')
                    content = row.get('题目内容', '')
                    answer = clean_answer(row.get('答案', ''), q_type)
                    
                    # 构建选项
                    options = {k: '' for k in ['A','B','C','D','E','F']}
                    if q_type == '判断题':
                        options.update({'A':'正确', 'B':'错误'})
                    else:
                        for opt in options:
                            options[opt] = str(row.get(opt, '')).strip()
                    
                    # 数据校验
                    if not answer:
                        raise ValueError("答案解析失败")
                    if not content:
                        raise ValueError("题干内容为空")
                        
                    # 添加结果
                    final_df.loc[len(final_df)] = {
                        '题目内容': content,
                        '答案': answer,
                        **options
                    }
                    
                except Exception as e:
                    error_log.append(f"[{file}] 第{idx+2}行错误：{str(e)} | 原始答案：{row.get('答案','')}")
                    
        except Exception as e:
            error_log.append(f"[{file}] 文件读取失败：{str(e)}")

    # 保存结果
    final_df.to_excel(output_path, index=False)
    
    # 生成日志
    if error_log:
        log_content = [
            f"处理时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"总处理文件数：{len(file_list)}",
            f"成功处理题目数：{len(final_df)}",
            f"发现错误数：{len(error_log)}",
            "\n错误详情：\n" + "\n".join(error_log)
        ]
        with open(log_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(log_content))
        print(f"处理完成，发现{len(error_log)}个问题，详见日志：{log_file}")
    else:
        print(f"完美处理！共转换{len(final_df)}道题目，输出文件：{output_path}")

if __name__ == "__main__":
    process_excel_files()