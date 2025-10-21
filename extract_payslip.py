#!/usr/bin/env python3
"""
Payslip Payment Details Extractor
提取工资单中的payment详情

重构版本 - 采用更清晰的模块化结构
"""

import re
import pdfplumber
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass


# ============================================================================
# 数据模型 (Data Models)
# ============================================================================

@dataclass
class PayPeriodInfo:
    """工资周期信息"""
    period: Optional[str] = None
    paid_date: Optional[str] = None


@dataclass
class PaymentRecord:
    """单条payment记录"""
    pdf_file: str
    page: int
    pay_period: Optional[str]
    paid_date: Optional[str]
    work_date: str
    hours: float
    rate: float
    amount: float

    def to_dict(self) -> Dict:
        """转换为字典格式"""
        return {
            'PDF File': self.pdf_file,
            'Page': self.page,
            'Pay Period': self.pay_period,
            'Paid Date': self.paid_date,
            'Work Date': self.work_date,
            'Hours': self.hours,
            'Rate': self.rate,
            'Amount': self.amount
        }


@dataclass
class SummaryRecord:
    """汇总记录"""
    pdf_file: str
    page: int
    pay_period: Optional[str]
    paid_date: Optional[str]
    gross_pay: Optional[float] = None
    tax: Optional[float] = None
    nett_pay: Optional[float] = None
    ytd_gross_pay: Optional[float] = None
    ytd_tax: Optional[float] = None
    ytd_nett_pay: Optional[float] = None
    disbursement_amount: Optional[float] = None

    def to_dict(self) -> Dict:
        """转换为字典格式"""
        return {
            'PDF File': self.pdf_file,
            'Page': self.page,
            'Pay Period': self.pay_period,
            'Paid Date': self.paid_date,
            'Gross Pay': self.gross_pay,
            'Tax': self.tax,
            'Nett Pay': self.nett_pay,
            'YTD Gross Pay': self.ytd_gross_pay,
            'YTD Tax': self.ytd_tax,
            'YTD Nett Pay': self.ytd_nett_pay,
            'Disbursement Amount': self.disbursement_amount
        }


# ============================================================================
# 配置和模式 (Configuration & Patterns)
# ============================================================================

class PatternConfig:
    """正则表达式模式配置"""

    # Pay Period和Paid Date模式
    PAY_PERIOD_PATTERN = r'Pay Period (.+?) to (.+?) Paid (.+?)$'

    # Payment行模式
    PAYMENT_PATTERN = r'CAS OrdPay \(incCASloading\)\s+([\d.]+)\s+([\d.]+)\s+(\w+)\s+([\d.]+)'

    # Summary模式
    GROSS_PAY_PATTERN = r'Gross Pay\s+([\d.]+)\s+([\d.]+)'
    TAX_PATTERN = r'^Tax\s+([\d.]+)\s+([\d.]+)'
    NETT_PAY_PATTERN = r'Nett Pay\s+([\d.]+)\s+([\d.]+)'
    DISBURSEMENT_PATTERN = r'Commonwealth Bank of Australia\s+\d+\s+([\d.]+)'

    # 日期格式
    REFERENCE_DATE_FORMAT = '%d%b%y'  # 例如: 01Jul24
    OUTPUT_DATE_FORMAT = '%Y-%m-%d'


# ============================================================================
# 工具类 (Utility Classes)
# ============================================================================

class DateParser:
    """日期解析器"""

    @staticmethod
    def parse_reference_date(ref_date: str) -> str:
        """
        解析reference日期格式

        Args:
            ref_date: 原始日期字符串 (如 01Jul24)

        Returns:
            格式化后的日期字符串 (如 2024-07-01)
        """
        try:
            date_obj = datetime.strptime(ref_date, PatternConfig.REFERENCE_DATE_FORMAT)
            return date_obj.strftime(PatternConfig.OUTPUT_DATE_FORMAT)
        except ValueError:
            return ref_date


# ============================================================================
# 文本解析器 (Text Parsers)
# ============================================================================

class PayPeriodParser:
    """工资周期解析器"""

    @staticmethod
    def extract(lines: List[str]) -> PayPeriodInfo:
        """
        从文本行中提取pay period信息

        Args:
            lines: 文本行列表

        Returns:
            PayPeriodInfo对象
        """
        for line in lines:
            if 'Pay Period' in line and 'Paid' in line:
                match = re.search(PatternConfig.PAY_PERIOD_PATTERN, line)
                if match:
                    period = f"{match.group(1)} to {match.group(2)}"
                    paid_date = match.group(3)
                    return PayPeriodInfo(period=period, paid_date=paid_date)

        return PayPeriodInfo()


class PaymentParser:
    """Payment条目解析器"""

    @staticmethod
    def extract(lines: List[str], pdf_filename: str, page_num: int,
                pay_period_info: PayPeriodInfo) -> List[PaymentRecord]:
        """
        从文本行中提取payment记录

        Args:
            lines: 文本行列表
            pdf_filename: PDF文件名
            page_num: 页码
            pay_period_info: 工资周期信息

        Returns:
            PaymentRecord列表
        """
        payments = []
        in_payments_section = False

        for i, line in enumerate(lines):
            # 检测Payments section开始
            if line.strip().startswith('Payments'):
                # "Hours"可能在同一行或下一行
                has_hours = ('Hours' in line or
                           (i + 1 < len(lines) and 'Hours' in lines[i + 1]))
                if has_hours:
                    in_payments_section = True
                    continue

            # 检测Payments section结束
            if in_payments_section:
                if (line.strip().startswith('Deductions') or
                    line.strip().startswith('Benefits')):
                    break

                # 尝试匹配payment行
                match = re.match(PatternConfig.PAYMENT_PATTERN, line)
                if match:
                    hours = float(match.group(1))
                    rate = float(match.group(2))
                    reference_date = match.group(3)
                    amount = float(match.group(4))

                    payment = PaymentRecord(
                        pdf_file=pdf_filename,
                        page=page_num,
                        pay_period=pay_period_info.period,
                        paid_date=pay_period_info.paid_date,
                        work_date=DateParser.parse_reference_date(reference_date),
                        hours=hours,
                        rate=rate,
                        amount=amount
                    )
                    payments.append(payment)

        return payments


class SummaryParser:
    """Summary信息解析器"""

    @staticmethod
    def extract(lines: List[str], pdf_filename: str, page_num: int,
                pay_period_info: PayPeriodInfo) -> SummaryRecord:
        """
        从文本行中提取summary信息

        Args:
            lines: 文本行列表
            pdf_filename: PDF文件名
            page_num: 页码
            pay_period_info: 工资周期信息

        Returns:
            SummaryRecord对象
        """
        summary = SummaryRecord(
            pdf_file=pdf_filename,
            page=page_num,
            pay_period=pay_period_info.period,
            paid_date=pay_period_info.paid_date
        )

        for line in lines:
            line_stripped = line.strip()

            # 提取Gross Pay
            if line_stripped.startswith('Gross Pay'):
                match = re.search(PatternConfig.GROSS_PAY_PATTERN, line)
                if match:
                    summary.gross_pay = float(match.group(1))
                    summary.ytd_gross_pay = float(match.group(2))

            # 提取Tax
            elif line_stripped.startswith('Tax') and not line.startswith('Tax'):
                match = re.search(PatternConfig.TAX_PATTERN, line)
                if match:
                    summary.tax = float(match.group(1))
                    summary.ytd_tax = float(match.group(2))

            # 提取Nett Pay
            elif line_stripped.startswith('Nett Pay'):
                match = re.search(PatternConfig.NETT_PAY_PATTERN, line)
                if match:
                    summary.nett_pay = float(match.group(1))
                    summary.ytd_nett_pay = float(match.group(2))

            # 提取Disbursement
            elif 'Commonwealth Bank of Australia' in line:
                match = re.search(PatternConfig.DISBURSEMENT_PATTERN, line)
                if match:
                    summary.disbursement_amount = float(match.group(1))

        return summary


# ============================================================================
# PDF处理器 (PDF Processor)
# ============================================================================

class PDFProcessor:
    """PDF文件处理器"""

    def __init__(self, pdf_path: str, pdf_filename: Optional[str] = None):
        """
        初始化PDF处理器

        Args:
            pdf_path: PDF文件路径
            pdf_filename: PDF文件名 (可选)
        """
        self.pdf_path = pdf_path
        self.pdf_filename = pdf_filename or Path(pdf_path).name
        self.payments: List[PaymentRecord] = []
        self.summaries: List[SummaryRecord] = []

    def process(self) -> Tuple[List[PaymentRecord], List[SummaryRecord]]:
        """
        处理PDF文件，提取所有数据

        Returns:
            (payments列表, summaries列表)元组
        """
        with pdfplumber.open(self.pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if text:
                    self._process_page(text, page_num)

        return self.payments, self.summaries

    def _process_page(self, text: str, page_num: int) -> None:
        """
        处理单个页面

        Args:
            text: 页面文本
            page_num: 页码
        """
        lines = text.split('\n')

        # 提取pay period信息
        pay_period_info = PayPeriodParser.extract(lines)

        # 提取payments
        payments = PaymentParser.extract(
            lines, self.pdf_filename, page_num, pay_period_info
        )
        self.payments.extend(payments)

        # 提取summary
        summary = SummaryParser.extract(
            lines, self.pdf_filename, page_num, pay_period_info
        )
        self.summaries.append(summary)


# ============================================================================
# 数据导出器 (Data Exporter)
# ============================================================================

class DataExporter:
    """数据导出器"""

    @staticmethod
    def to_dataframes(payments: List[PaymentRecord],
                     summaries: List[SummaryRecord]) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        将记录列表转换为DataFrame

        Args:
            payments: Payment记录列表
            summaries: Summary记录列表

        Returns:
            (payments_df, summaries_df)元组
        """
        payments_data = [p.to_dict() for p in payments]
        summaries_data = [s.to_dict() for s in summaries]

        payments_df = pd.DataFrame(payments_data)
        summaries_df = pd.DataFrame(summaries_data)

        return payments_df, summaries_df

    @staticmethod
    def save_to_excel(payments_df: pd.DataFrame, summaries_df: pd.DataFrame,
                     output_file: str = 'payslip_details.xlsx') -> None:
        """
        保存数据到Excel文件

        Args:
            payments_df: Payments DataFrame
            summaries_df: Summaries DataFrame
            output_file: 输出文件名
        """
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            payments_df.to_excel(writer, sheet_name='Payment Details', index=False)
            summaries_df.to_excel(writer, sheet_name='Summary', index=False)

    @staticmethod
    def print_statistics(payments_df: pd.DataFrame, summaries_df: pd.DataFrame) -> None:
        """
        打印统计信息

        Args:
            payments_df: Payments DataFrame
            summaries_df: Summaries DataFrame
        """
        print("\n=== 统计信息 ===")
        print(f"总工作小时数: {payments_df['Hours'].sum():.2f}")
        print(f"总收入 (Gross): ${summaries_df['Gross Pay'].sum():,.2f}")
        print(f"总税额: ${summaries_df['Tax'].sum():,.2f}")
        print(f"总净收入 (Nett): ${summaries_df['Nett Pay'].sum():,.2f}")

    @staticmethod
    def print_sample_data(payments_df: pd.DataFrame, n: int = 10) -> None:
        """
        打印样本数据

        Args:
            payments_df: Payments DataFrame
            n: 显示的行数
        """
        print(f"\n=== 前{n}条Payment记录 ===")
        print(payments_df.head(n).to_string(index=False))


# ============================================================================
# 主处理器 (Main Processor)
# ============================================================================

class PayslipProcessor:
    """工资单处理器主类"""

    def __init__(self, pdf_directory: Path = Path('.')):
        """
        初始化工资单处理器

        Args:
            pdf_directory: PDF文件所在目录
        """
        self.pdf_directory = pdf_directory
        self.all_payments: List[PaymentRecord] = []
        self.all_summaries: List[SummaryRecord] = []

    def find_pdf_files(self) -> List[Path]:
        """
        查找目录下的所有PDF文件

        Returns:
            PDF文件路径列表
        """
        return list(self.pdf_directory.glob('data/*.pdf'))

    def process_all_pdfs(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        处理所有PDF文件

        Returns:
            (payments_df, summaries_df)元组
        """
        pdf_files = self.find_pdf_files()

        if not pdf_files:
            print("未找到PDF文件！")
            return pd.DataFrame(), pd.DataFrame()

        print(f"找到 {len(pdf_files)} 个PDF文件\n")

        # 处理每个PDF文件
        for pdf_file in pdf_files:
            self._process_single_pdf(pdf_file)

        # 转换为DataFrame
        if self.all_payments:
            return DataExporter.to_dataframes(self.all_payments, self.all_summaries)
        else:
            print("未提取到任何数据！")
            return pd.DataFrame(), pd.DataFrame()

    def _process_single_pdf(self, pdf_file: Path) -> None:
        """
        处理单个PDF文件

        Args:
            pdf_file: PDF文件路径
        """
        print(f"正在处理: {pdf_file.name}")

        processor = PDFProcessor(str(pdf_file), pdf_file.name)
        payments, summaries = processor.process()

        self.all_payments.extend(payments)
        self.all_summaries.extend(summaries)

        print(f"  - 提取了 {len(payments)} 条payment记录")
        print(f"  - 提取了 {len(summaries)} 个pay period汇总\n")


# ============================================================================
# 主函数 (Main Function)
# ============================================================================

def main():
    """主函数 - 协调整个提取流程"""
    # 创建处理器
    processor = PayslipProcessor()

    # 处理所有PDF
    payments_df, summaries_df = processor.process_all_pdfs()

    # 如果有数据，保存并显示
    if not payments_df.empty:
        # 保存到Excel
        output_file = 'payslip_details.xlsx'
        DataExporter.save_to_excel(payments_df, summaries_df, output_file)

        print(f"✓ 数据已保存到: {output_file}")
        print(f"  - Payment Details: {len(payments_df)} 条记录")
        print(f"  - Summary: {len(summaries_df)} 条记录")

        # 显示统计信息
        DataExporter.print_statistics(payments_df, summaries_df)

        # 显示样本数据
        DataExporter.print_sample_data(payments_df)

        return payments_df, summaries_df


if __name__ == '__main__':
    main()
