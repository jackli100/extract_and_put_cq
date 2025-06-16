import pandas as pd
import numpy as np
from datetime import datetime

def read_f_column_and_calculate_sum():
    """
    读取0615.xlsx文件的F列数据，统计和值
    当无法读取数据时报错并打印该行的全部信息
    """
    file_name = "0615.xlsx"
    
    print("🚀 F列数据统计程序启动")
    print(f"⏰ 执行时间: 2025-06-15 09:28:14 UTC")
    print(f"👤 执行用户: jackli100")
    print(f"📁 目标文件: {file_name}")
    print("=" * 60)
    
    try:
        # 读取Excel文件，包含首行标题在内
        print("📖 正在读取Excel文件...")
        df = pd.read_excel(file_name, header=None)
        print(f"✅ 文件读取成功，共 {len(df)} 行数据")
        
        # 检查是否有F列（索引为5，因为从0开始计数）
        col_F_index = 5  # F列是第6列，索引为5
        
        if col_F_index >= len(df.columns):
            print(f"❌ 错误：文件中没有F列（第6列），实际只有 {len(df.columns)} 列")
            print("可用的列:", list(df.columns))
            return
        
        # 获取F列数据
        f_column_data = df.iloc[:, col_F_index]
        column_name = df.columns[col_F_index] if col_F_index < len(df.columns) else f"第{col_F_index+1}列"
        
        print(f"\n📊 开始分析F列数据")
        print(f"列名: {column_name}")
        print(f"数据类型: {f_column_data.dtype}")
        print(f"总行数: {len(f_column_data)}")
        
        # 统计和处理数据
        valid_sum = 0
        valid_count = 0
        error_count = 0
        error_details = []
        
        print(f"\n🔍 逐行检查F列数据...")
        print("-" * 80)
        
        for index, value in f_column_data.items():
            excel_row = index + 1  # Excel行号从1开始
            
            try:
                # 尝试转换为浮点数
                if pd.isna(value) or value == '' or str(value).strip() == '':
                    # 空值情况
                    print(f"⚠️  第 {excel_row} 行 F列数据为空值")
                    error_count += 1
                    error_details.append({
                        'excel_row': excel_row,
                        'pandas_index': index,
                        'error_type': '空值',
                        'f_column_value': value
                    })
                    # 打印该行的全部信息
                    print(f"   该行完整数据: {dict(df.iloc[index])}")
                    print("-" * 80)
                else:
                    # 尝试转换为数值
                    clean_value = str(value).strip()
                    numeric_value = float(clean_value)
                    valid_sum += numeric_value
                    valid_count += 1
                    print(f"✅ 第 {excel_row} 行 F列: {numeric_value:.2f}")
                    
            except (ValueError, TypeError) as e:
                # 转换失败
                error_count += 1
                print(f"❌ 第 {excel_row} 行 F列数据转换失败: '{value}'")
                print(f"   错误原因: {str(e)}")
                
                error_details.append({
                    'excel_row': excel_row,
                    'pandas_index': index,
                    'error_type': f'转换错误: {str(e)}',
                    'f_column_value': value
                })
                
                # 打印该行的全部信息
                row_data = df.iloc[index]
                print(f"   该行完整数据:")
                for col_name, col_value in row_data.items():
                    print(f"     {col_name}: {col_value}")
                print("-" * 80)
                
            except Exception as e:
                # 其他未预期的错误
                error_count += 1
                print(f"💥 第 {excel_row} 行处理时发生未知错误: {str(e)}")
                
                error_details.append({
                    'excel_row': excel_row,
                    'pandas_index': index,
                    'error_type': f'未知错误: {str(e)}',
                    'f_column_value': value
                })
                
                # 打印该行的全部信息
                try:
                    row_data = df.iloc[index]
                    print(f"   该行完整数据:")
                    for col_name, col_value in row_data.items():
                        print(f"     {col_name}: {col_value}")
                except:
                    print(f"   无法获取该行完整数据")
                print("-" * 80)
        
        # 输出最终统计结果
        print(f"\n" + "📈 F列数据统计结果".center(60, "="))
        print(f"📊 数据概览:")
        print(f"  总行数: {len(f_column_data)}")
        print(f"  有效数据行数: {valid_count}")
        print(f"  无效数据行数: {error_count}")
        print(f"  数据有效率: {(valid_count/len(f_column_data)*100):.1f}%")
        
        print(f"\n💰 F列数值统计:")
        print(f"  F列数据总和: {valid_sum:.2f}")
        
        if valid_count > 0:
            # 计算有效数据的统计信息
            valid_data = []
            for index, value in f_column_data.items():
                try:
                    if not (pd.isna(value) or value == '' or str(value).strip() == ''):
                        clean_value = str(value).strip()
                        numeric_value = float(clean_value)
                        valid_data.append(numeric_value)
                except:
                    pass
            
            if valid_data:
                print(f"  最小值: {min(valid_data):.2f}")
                print(f"  最大值: {max(valid_data):.2f}")
                print(f"  平均值: {sum(valid_data)/len(valid_data):.2f}")
        
        # 错误详情汇总
        if error_details:
            print(f"\n⚠️  错误详情汇总:")
            error_types = {}
            for detail in error_details:
                error_type = detail['error_type']
                if error_type not in error_types:
                    error_types[error_type] = []
                error_types[error_type].append(detail['excel_row'])
            
            for error_type, rows in error_types.items():
                print(f"  {error_type}: {len(rows)} 行")
                print(f"    Excel行号: {', '.join(map(str, sorted(rows)))}")
        else:
            print(f"\n✅ 所有F列数据都是有效的数值型数据")
        
        print(f"\n🎯 最终结果: F列数据总和 = {valid_sum:.2f}")
        
        return valid_sum, valid_count, error_count
        
    except FileNotFoundError:
        print(f"❌ 文件未找到: {file_name}")
        print("请确保文件 '0615.xlsx' 存在于当前目录中")
        return None, None, None
        
    except pd.errors.EmptyDataError:
        print(f"❌ 文件 '{file_name}' 是空文件或无法读取")
        return None, None, None
        
    except Exception as e:
        print(f"❌ 读取文件时发生错误: {str(e)}")
        import traceback
        print("\n详细错误信息:")
        traceback.print_exc()
        return None, None, None

def main():
    """
    主函数
    """
    print("=" * 60)
    result = read_f_column_and_calculate_sum()
    
    if result[0] is not None:
        total_sum, valid_count, error_count = result
        print("\n" + "🏁 程序执行完毕".center(60, "="))
        print(f"✅ 成功处理 F列数据")
        print(f"📊 最终统计: 总和={total_sum:.2f}, 有效行={valid_count}, 错误行={error_count}")
    else:
        print("\n" + "💥 程序执行失败".center(60, "="))
        print(f"❌ 无法完成 F列数据统计")
    
    print(f"⏰ 程序结束时间: 2025-06-15 09:28:14 UTC")

if __name__ == "__main__":
    main()

