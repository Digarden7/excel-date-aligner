# -*- coding: utf-8 -*-
"""
Excel 日期对齐合并工具
======================
按日期列对齐合并 Excel 多工作表数据，支持缺失值前向填充。

使用方法：
    python excel_date_aligner.py <输入文件> [--output <输出文件>] [--method ffill|bfill]

示例：
    python excel_date_aligner.py data.xlsx
    python excel_date_aligner.py data.xlsx --output result.xlsx
    python excel_date_aligner.py data.xlsx --method bfill
"""

import pandas as pd
import argparse
import os


def align_excel(filepath, output=None, method='ffill'):
    """
    按日期对齐合并 Excel 多工作表数据。

    参数：
        filepath (str): 输入 Excel 文件路径
        output (str): 输出文件路径，默认在输入文件同目录下生成
        method (str): 缺失值填充方式，'ffill' 前向填充，'bfill' 后向填充

    返回：
        pd.DataFrame: 合并后的数据
    """
    # 读取所有工作表
    xls = pd.ExcelFile(filepath)
    sheet_names = xls.sheet_names

    if len(sheet_names) < 2:
        print(f"警告：文件仅包含 {len(sheet_names)} 个工作表，无需合并")
        return pd.read_excel(filepath, sheet_name=0)

    # 读取前两个工作表
    df1 = pd.read_excel(filepath, sheet_name=0)
    df2 = pd.read_excel(filepath, sheet_name=1)

    # 按日期左连接合并（以第一个工作表为基准）
    dff = pd.merge(df1, df2, how='left', on='日期')

    # 如果有更多工作表，继续合并
    for i in range(2, len(sheet_names)):
        dfn = pd.read_excel(filepath, sheet_name=i)
        if '日期' in dfn.columns:
            dff = pd.merge(dff, dfn, how='left', on='日期')

    # 缺失值填充
    if method == 'ffill':
        dff = dff.ffill(axis=0)
        print('已使用前向填充 (ffill) 补齐缺失值')
    elif method == 'bfill':
        dff = dff.bfill(axis=0)
        print('已使用后向填充 (bfill) 补齐缺失值')

    # 确定输出路径
    if output is None:
        base, ext = os.path.splitext(filepath)
        output = f'{base}_对齐后{ext}'

    # 保存结果
    dff.to_excel(output, index=False)
    print(f'合并完成！输出文件：{output}')
    print(f'合并后数据：{dff.shape[0]} 行 x {dff.shape[1]} 列')

    return dff


def main():
    parser = argparse.ArgumentParser(
        description='按日期对齐合并 Excel 多工作表数据，支持缺失值前向填充'
    )
    parser.add_argument('input', help='输入 Excel 文件路径')
    parser.add_argument('-o', '--output', help='输出文件路径（默认自动生成）')
    parser.add_argument(
        '-m', '--method',
        choices=['ffill', 'bfill'],
        default='ffill',
        help='缺失值填充方式：ffill=前向填充（默认），bfill=后向填充'
    )
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f'错误：文件不存在 - {args.input}')
        return

    align_excel(args.input, args.output, args.method)


if __name__ == '__main__':
    main()
