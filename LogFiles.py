import re
from datetime import datetime
from collections import defaultdict

class LogAnalyzer:
    def __init__(self, log_file_path):
        self.log_file_path = log_file_path
        self.send_sequences = []    # 发送数据的通信序列
        self.receive_sequences = [] # 接收数据的通信序列
        
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
    def print_sequence_details(self, sequence, seq_type, idx):
        if len(sequence) >= 4:  # 确保至少包含完整的通信过程
            start_time = sequence[0]['timestamp']
            end_time = sequence[-1]['timestamp']
            duration = (end_time - start_time).total_seconds() * 1000
            
            print(f"\n{seq_type} 序列 {idx}:")
            print(f"开始时间: {start_time}")
            print(f"结束时间: {end_time}")
            print(f"持续时间: {duration:.2f}ms")
            print("通信内容:")
            
            main_data = None
            for msg in sequence:
                print(f"  {msg['timestamp'].strftime('%H:%M:%S.%f')[:-3]} "
                        f"{msg['direction']}: {msg['data']}")
                
                # 提取主报文（非05/04/06的数据）
                if ((seq_type == "发送" and msg['direction'] == 'Snd') or 
                    (seq_type == "接收" and msg['direction'] == 'Rcv')) and \
                    msg['data'] not in ['05', '04', '06']:
                    main_data = msg['data']
            
            if main_data:
                print(f"主报文内容: {main_data}")
            print("-" * 40)
    def print_statistics(self):
        print("\n=== 通信序列分析 ===")
        print("\n发送数据的通信序列:")
        print("=" * 80)
        for idx, sequence in enumerate(self.send_sequences, 1):
            self.print_sequence_details(sequence, "发送", idx)
            
        print("\n接收数据的通信序列:")
        print("=" * 80)
        for idx, sequence in enumerate(self.receive_sequences, 1):
            self.print_sequence_details(sequence, "接收", idx)
        
        # 打印总体统计
        print("\n总体统计:")
        print(f"发送序列总数: {len(self.send_sequences)}")
        print(f"接收序列总数: {len(self.receive_sequences)}")

def main():
    analyzer = LogAnalyzer('F:/01_ProjectsFiles/01_Programs/LogFilse/2024-09-24.log')
    analyzer.analyze_log()
    analyzer.print_statistics()
if __name__ == "__main__":
    main()