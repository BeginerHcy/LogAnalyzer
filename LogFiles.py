import re
from datetime import datetime
from collections import defaultdict
import pandas as pd
import os

rootpath = 'D:/00_PROJECT/04_Python/LogAnalyzer/'

def inser_char(s,char,index,len):
    return s[:index] + char[0:len] + s[index:]
    
class LogAnalyzer:
    def __init__(self, log_file_path):
        self.log_file_path = log_file_path
        self.send_sequences = []    
        self.receive_sequences = [] 
        # 用于存储Excel数据的列表
        self.excel_data = {
            'send': [],
            'receive': []
        }
    def parse_log_line(self, line):
        pattern = r'Debug:\s+(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.(\d{3}|\d{1}|\d{2})):\s+(Snd|Rcv):\s+(.+)'
        match = re.match(pattern, line)
        #print(match.group(1))
        if match:
            fslen = len(match.group(1).split('.')[-1])
            newgroup1 = (inser_char(match.group(1),'000',-fslen, 3-fslen))
            timestamp = datetime.strptime(newgroup1, '%Y-%m-%d %H:%M:%S.%f')
            #print(timestamp)
            direction = match.group(3)
            data = match.group(4).strip()
            return {
                'timestamp': timestamp,
                'direction': direction,
                'data': data
            }
        return None
    def analyze_log(self):
        with open(self.log_file_path, 'r') as f:
            current_send_sequence = []
            current_receive_sequence = []
            
            for line in f:
                parsed = self.parse_log_line(line.strip())
                #print(parsed)
                if not parsed:
                    continue
                
                # 处理发送数据的通信序列
                if parsed['direction'] == 'Snd' and parsed['data'] == '05':
                    if current_send_sequence:
                        self.send_sequences.append(current_send_sequence)
                    current_send_sequence = [parsed]
                    continue
                    
                # 处理接收数据的通信序列
                if parsed['direction'] == 'Rcv' and parsed['data'] == '05':
                    if current_receive_sequence:
                        self.receive_sequences.append(current_receive_sequence)
                    current_receive_sequence = [parsed]
                    continue
                    # 更新发送序列
                if current_send_sequence:
                    current_send_sequence.append(parsed)
                    if parsed['direction'] == 'Rcv' and parsed['data'] == '06':
                        self.send_sequences.append(current_send_sequence)
                        current_send_sequence = []
                
                # 更新接收序列
                if current_receive_sequence:
                    current_receive_sequence.append(parsed)
                    if parsed['direction'] == 'Snd' and parsed['data'] == '06':
                        self.receive_sequences.append(current_receive_sequence)
                        current_receive_sequence = []
    def sequence_to_excel_row(self, sequence, seq_type, idx):
        if len(sequence) >= 4:  # 确保至少包含完整的通信过程
            start_time = sequence[0]['timestamp']
            end_time = sequence[-1]['timestamp']
            duration = (end_time - start_time).total_seconds() * 1000
            
            # 提取主体报文
            main_data = []
            found_04 = False
            direction04 = 'Rcv'
            for msg in sequence:
                if found_04 and msg['data'] != '06':  # 在找到04之后，06之前的都是主体报文
                    main_data.append(msg['data'])
                if msg['data'] == '04':
                    found_04 = True
                    direction04 = msg['direction']
                elif msg['data'] == '06' and msg['direction'] == direction04 :
                    break
            
            # 构建完整的通信过程字符串
            sequence_str = []
            for msg in sequence:
                sequence_str.append(f"{msg['timestamp'].strftime('%H:%M:%S.%f')[:-3]} "
                                    f"{msg['direction']}: {msg['data']}")
            return {
                '序列号': idx,
                '开始时间': start_time,
                '结束时间': end_time,
                '持续时间(ms)': round(duration, 2),
                '主报文': ' '.join(main_data) if main_data else '',
                '完整通信过程': '\n'.join(sequence_str)
            }
        return None
    def generate_excel_report(self):
        # 准备所有序列的数据
        all_sequences = []
        
        # 处理发送序列
        for idx, sequence in enumerate(self.send_sequences, 1):
            row_data = self.sequence_to_excel_row(sequence, "发送", idx)
            if row_data:
                row_data['类型'] = '发送'  # 添加类型标识
                all_sequences.append(row_data)
        
        # 处理接收序列
        for idx, sequence in enumerate(self.receive_sequences, 1):
            row_data = self.sequence_to_excel_row(sequence, "接收", idx)
            if row_data:
                row_data['类型'] = '接收'  # 添加类型标识
                all_sequences.append(row_data)
        
        # 按开始时间排序
        all_sequences.sort(key=lambda x: x['开始时间'])
        
        # 重新编号
        for idx, sequence in enumerate(all_sequences, 1):
            sequence['序列号'] = idx
        
        # 创建输出目录
        output_dir = rootpath + 'analysis_results'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 生成文件名
        current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_path = f'{output_dir}/log_analysis_{current_time}.xlsx'
        
        # 创建Excel写入器
        with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
            # 所有数据序列表
            if all_sequences:
                #print(all_sequences)
                df_all = pd.DataFrame(all_sequences)
                #print(df_all)
                # 重新排列列顺序，将'类型'列放在序列号后面
                columns = ['序列号', '类型', '开始时间', '结束时间', '持续时间(ms)', '主报文', '完整通信过程']
                df_all = df_all[columns]
                df_all.to_excel(writer, sheet_name='通信序列', index=False)
                # 调整列宽
                worksheet = writer.sheets['通信序列']
                worksheet.set_column('A:A', 8)   # 序列号
                worksheet.set_column('B:B', 8)   # 类型
                worksheet.set_column('C:D', 20)  # 时间列
                worksheet.set_column('E:E', 12)  # 持续时间
                worksheet.set_column('F:F', 30)  # 主报文
                worksheet.set_column('G:G', 50)  # 完整通信过程
                # 添加条件格式，为不同类型设置不同的背景色
                format_send = writer.book.add_format({'bg_color': '#E6F3FF'})  # 淡蓝色
                format_receive = writer.book.add_format({'bg_color': '#F3FFE6'})  # 淡绿色
                # 应用条件格式
                worksheet.conditional_format('A2:G1048576', {
                    'type': 'formula',
                    'criteria': '=$B2="发送"',
                    'format': format_send
                })
                worksheet.conditional_format('A2:G1048576', {
                    'type': 'formula',
                    'criteria': '=$B2="接收"',
                    'format': format_receive
                })
                # 添加统计信息表
                #print('works here')
                stats_data = {
                '统计项': ['发送序列总数', '接收序列总数', '总序列数'],
                '数量': [
                    len(self.send_sequences), 
                    len(self.receive_sequences),
                    len(self.send_sequences) + len(self.receive_sequences)
                    ]
                }
                #print(stats_data)
            df_stats = pd.DataFrame(stats_data)
            df_stats.to_excel(writer, sheet_name='统计信息', index=False)
            
            # 调整统计信息表的列宽
            worksheet = writer.sheets['统计信息']
            worksheet.set_column('A:A', 15)
            worksheet.set_column('B:B', 10)
        print(f"\nExcel报告已生成: {excel_path}")
    def print_statistics(self):
        # ... 保持原有的 print_statistics 方法不变 ...
        # 添加生成Excel报告
        self.generate_excel_report()
def main():
    #analyzer = LogAnalyzer('F:/01_ProjectsFiles/01_Programs/LogFilse/2024-09-24.log')
    analyzer = LogAnalyzer(rootpath + 'Logcom.log')
    analyzer.analyze_log()
    analyzer.print_statistics()
if __name__ == "__main__":
    main()