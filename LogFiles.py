import re
from datetime import datetime
from collections import defaultdict
import pandas as pd
import os
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
        pattern = r'Debug:\s+(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.\d{3}):\s+(Snd|Rcv):\s+(.+)'
        match = re.match(pattern, line)
        if match:
            timestamp = datetime.strptime(match.group(1), '%Y-%m-%d %H:%M:%S.%f')
            direction = match.group(2)
            data = match.group(3).strip()
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
    # ... 保持 parse_log_line 和 analyze_log 方法不变 ...
    def sequence_to_excel_row(self, sequence, seq_type, idx):
        if len(sequence) >= 4:
            start_time = sequence[0]['timestamp']
            end_time = sequence[-1]['timestamp']
            duration = (end_time - start_time).total_seconds() * 1000
            
            main_data = None
            sequence_str = []
            for msg in sequence:
                sequence_str.append(f"{msg['timestamp'].strftime('%H:%M:%S.%f')[:-3]} "
                                    f"{msg['direction']}: {msg['data']}")
                
                if ((seq_type == "发送" and msg['direction'] == 'Snd') or 
                    (seq_type == "接收" and msg['direction'] == 'Rcv')) and \
                    msg['data'] not in ['05', '04', '06']:
                    main_data = msg['data']
            
            return {
                '序列号': idx,
                '开始时间': start_time,
                '结束时间': end_time,
                '持续时间(ms)': round(duration, 2),
                '主报文': main_data if main_data else '',
                '完整通信过程': '\n'.join(sequence_str)
                }
        return None
    def generate_excel_report(self):
        # 准备Excel数据
        for idx, sequence in enumerate(self.send_sequences, 1):
            row_data = self.sequence_to_excel_row(sequence, "发送", idx)
            if row_data:
                self.excel_data['send'].append(row_data)
        
        for idx, sequence in enumerate(self.receive_sequences, 1):
            row_data = self.sequence_to_excel_row(sequence, "接收", idx)
            if row_data:
                self.excel_data['receive'].append(row_data)
            # 创建Excel文件
        output_dir = 'analysis_results'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            # 生成文件名（使用当前时间）
        current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_path = f'{output_dir}/log_analysis_{current_time}.xlsx'
            # 创建Excel写入器
        with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
            # 发送数据表
            if self.excel_data['send']:
                df_send = pd.DataFrame(self.excel_data['send'])
                df_send.to_excel(writer, sheet_name='发送数据序列', index=False)
                
                # 调整列宽
                worksheet = writer.sheets['发送数据序列']
                worksheet.set_column('A:A', 8)  # 序列号
                worksheet.set_column('B:C', 20)  # 时间列
                worksheet.set_column('D:D', 12)  # 持续时间
                worksheet.set_column('E:E', 30)  # 主报文
                worksheet.set_column('F:F', 50)  # 完整通信过程
                # 接收数据表
            if self.excel_data['receive']:
                df_receive = pd.DataFrame(self.excel_data['receive'])
                df_receive.to_excel(writer, sheet_name='接收数据序列', index=False)
                
                # 调整列宽
                worksheet = writer.sheets['接收数据序列']
                worksheet.set_column('A:A', 8)
                worksheet.set_column('B:C', 20)
                worksheet.set_column('D:D', 12)
                worksheet.set_column('E:E', 30)
                worksheet.set_column('F:F', 50)
                # 添加统计信息表
            stats_data = {
                '统计项': ['发送序列总数', '接收序列总数'],
                '数量': [len(self.send_sequences), len(self.receive_sequences)]
            }
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
    analyzer = LogAnalyzer('F:/01_ProjectsFiles/01_Programs/LogFilse/2024-09-24.log')
    analyzer.analyze_log()
    analyzer.print_statistics()
if __name__ == "__main__":
    main()