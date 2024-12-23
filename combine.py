import re
from datetime import datetime
import pandas as pd
import os
class CombinedLogAnalyzer:
    def __init__(self, comm_log_path, control_log_path):
        self.comm_log_path = comm_log_path
        self.control_log_path = control_log_path
        self.send_sequences = []    
        self.receive_sequences = []
        self.control_records = []  # 用于存储控制日志的记录
        
    # ... 保持原有的 parse_log_line, parse_main_data 等方法 ...
    def analyze_control_log(self):
        """分析控制日志"""
        with open(self.control_log_path, 'r') as f:
            # 这里添加控制日志的分析逻辑
            # 将结果存储在 self.control_records 中
            pass
    def generate_excel_report(self):
        # 准备通信序列数据
        all_sequences = []
        for idx, sequence in enumerate(self.send_sequences, 1):
            row_data = self.sequence_to_excel_row(sequence, "发送", idx)
            if row_data:
                row_data['类型'] = '发送'
                all_sequences.append(row_data)
        
        for idx, sequence in enumerate(self.receive_sequences, 1):
            row_data = self.sequence_to_excel_row(sequence, "接收", idx)
            if row_data:
                row_data['类型'] = '接收'
                all_sequences.append(row_data)
        
        # 按开始时间排序
        all_sequences.sort(key=lambda x: x['开始时间'])
        
        # 创建输出目录
        output_dir = 'analysis_results'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # 生成文件名
        current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_path = f'{output_dir}/combined_analysis_{current_time}.xlsx'
        
        # 创建Excel写入器
        with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
            # 通信序列表
            if all_sequences:
                df_comm = pd.DataFrame(all_sequences)
                columns = ['序列号', '类型', '开始时间', '结束时间', '持续时间(ms)', 
                            '主报文', '原始报文', '完整通信过程']
                df_comm = df_comm[columns]
                df_comm.to_excel(writer, sheet_name='通信序列', index=False)
                
                # 设置通信序列表格式
                worksheet = writer.sheets['通信序列']
                self._format_comm_sheet(worksheet, writer.book)
                # 控制日志表
            if self.control_records:
                df_control = pd.DataFrame(self.control_records)
                df_control.to_excel(writer, sheet_name='控制日志', index=False)
                
                # 设置控制日志表格式
                worksheet = writer.sheets['控制日志']
                self._format_control_sheet(worksheet)
                # 统计信息表
            stats_data = {
                '统计项': ['发送序列总数', '接收序列总数', '总序列数', '控制日志记录数'],
                '数量': [
                    len(self.send_sequences),
                    len(self.receive_sequences),
                    len(self.send_sequences) + len(self.receive_sequences),
                    len(self.control_records)
                ]
            }
            df_stats = pd.DataFrame(stats_data)
            df_stats.to_excel(writer, sheet_name='统计信息', index=False)
            
            # 设置统计信息表格式
            worksheet = writer.sheets['统计信息']
            worksheet.set_column('A:A', 15)
            worksheet.set_column('B:B', 10)
        print(f"\nExcel报告已生成: {excel_path}")
    def _format_comm_sheet(self, worksheet, workbook):
        """设置通信序列表的格式"""
        worksheet.set_column('A:A', 8)   # 序列号
        worksheet.set_column('B:B', 8)   # 类型
        worksheet.set_column('C:D', 20)  # 时间列
        worksheet.set_column('E:E', 12)  # 持续时间
        worksheet.set_column('F:F', 40)  # 主报文解析
        worksheet.set_column('G:G', 30)  # 原始报文
        worksheet.set_column('H:H', 50)  # 完整通信过程
        
        # 添加条件格式
        format_send = workbook.add_format({'bg_color': '#E6F3FF'})
        format_receive = workbook.add_format({'bg_color': '#F3FFE6'})
        
        worksheet.conditional_format('A2:H1048576', {
            'type': 'formula',
            'criteria': '=$B2="发送"',
            'format': format_send
        })
        worksheet.conditional_format('A2:H1048576', {
            'type': 'formula',
            'criteria': '=$B2="接收"',
            'format': format_receive
        })
    def _format_control_sheet(self, worksheet):
        """设置控制日志表的格式"""
        # 根据控制日志的具体列来设置格式
        pass
def main():
    analyzer = CombinedLogAnalyzer(
        'LogFiles/comm.log',
        'LogFiles/control.log'
    )
    analyzer.analyze_log()  # 分析通信日志
    analyzer.analyze_control_log()  # 分析控制日志
    analyzer.generate_excel_report()
if __name__ == "__main__":
    main()