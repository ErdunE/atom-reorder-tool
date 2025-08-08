#!/bin/bash

# 原子顺序恢复脚本 - 修正版
# 功能：将Excel中末尾的H原子按正确的循环模式重新排列
#
# 正确的循环模式：
# 第一个循环: O H C H H C H H (8原子)
# 中间循环: O C H H C H H (7原子)
# 最后循环: O C H H C H H H (8原子)

# 检查依赖
check_dependencies() {
    if ! command -v python3 &> /dev/null; then
        echo "错误: 需要安装 python3"
        exit 1
    fi
    
    # 检查Python库
    python3 -c "import pandas, openpyxl" 2>/dev/null
    if [ $? -ne 0 ]; then
        echo "错误: 需要安装 pandas 和 openpyxl"
        echo "请运行: pip3 install pandas openpyxl"
        exit 1
    fi
}

# 显示使用帮助
show_help() {
    echo "原子顺序恢复脚本 - 修正版"
    echo ""
    echo "用法: $0 [选项] <输入文件>"
    echo ""
    echo "选项:"
    echo "  -h, --help          显示此帮助信息"
    echo "  -o, --output FILE   指定输出文件名 (默认: input_reordered.xlsx)"
    echo "  -v, --verbose       详细输出"
    echo "  -p, --preview       只预览不保存"
    echo ""
    echo "示例:"
    echo "  $0 input.xlsx"
    echo "  $0 -o output.xlsx -v input.xlsx"
    echo ""
    echo "循环模式:"
    echo "  第一个循环: O H C H H C H H"
    echo "  中间循环:   O C H H C H H"
    echo "  最后循环:   O C H H C H H H"
}

# 默认参数
INPUT_FILE=""
OUTPUT_FILE=""
VERBOSE=false
PREVIEW_ONLY=false

# 解析命令行参数
while [[ $# -gt 0 ]]; do
    case $1 in
        -h|--help)
            show_help
            exit 0
            ;;
        -o|--output)
            OUTPUT_FILE="$2"
            shift 2
            ;;
        -v|--verbose)
            VERBOSE=true
            shift
            ;;
        -p|--preview)
            PREVIEW_ONLY=true
            shift
            ;;
        -*)
            echo "未知选项: $1"
            show_help
            exit 1
            ;;
        *)
            if [ -z "$INPUT_FILE" ]; then
                INPUT_FILE="$1"
            else
                echo "错误: 只能指定一个输入文件"
                exit 1
            fi
            shift
            ;;
    esac
done

# 检查输入文件
if [ -z "$INPUT_FILE" ]; then
    echo "错误: 请指定输入文件"
    show_help
    exit 1
fi

if [ ! -f "$INPUT_FILE" ]; then
    echo "错误: 文件 '$INPUT_FILE' 不存在"
    exit 1
fi

# 设置默认输出文件名
if [ -z "$OUTPUT_FILE" ]; then
    OUTPUT_FILE="${INPUT_FILE%.*}_reordered.xlsx"
fi

# 记录开始时间
START_TIME=$(date +%s)

echo "🧪 原子顺序恢复工具 - 修正版"
echo "================================"
echo "输入文件: $INPUT_FILE"
echo "输出文件: $OUTPUT_FILE"
echo "详细模式: $VERBOSE"
echo "预览模式: $PREVIEW_ONLY"
echo ""

# 检查依赖
if [ "$VERBOSE" = true ]; then
    echo "📋 检查依赖..."
fi
check_dependencies

# 创建Python脚本进行处理
PYTHON_SCRIPT=$(cat << 'EOF'
import pandas as pd
import sys
import os

def log_verbose(message, verbose=False):
    if verbose:
        print(message)

def reorder_atoms(input_file, output_file, verbose=False, preview_only=False):
    # 正确的循环模式定义
    FIRST_PATTERN = ['O', 'H', 'C', 'H', 'H', 'C', 'H', 'H']      # 第一个循环: 8原子
    MIDDLE_PATTERN = ['O', 'C', 'H', 'H', 'C', 'H', 'H']          # 中间循环: 7原子
    LAST_PATTERN = ['O', 'C', 'H', 'H', 'C', 'H', 'H', 'O', 'H']  # 最后循环: 8原子
    
    try:
        # 读取Excel文件
        log_verbose("📖 正在读取Excel文件...", verbose)
        df = pd.read_excel(input_file, header=None)
        log_verbose(f"   读取到 {len(df)} 行数据", verbose)
        
        # 分离H原子和非H原子
        h_atoms = []
        non_h_atoms = []
        
        for index, row in df.iterrows():
            atom_name = str(row[0]).strip()
            if atom_name.startswith('H(') and atom_name.endswith(')'):
                h_atoms.append((index, row))
            else:
                non_h_atoms.append((index, row))
        
        print(f"📊 数据分析:")
        print(f"   H原子数量: {len(h_atoms)}")
        print(f"   非H原子数量: {len(non_h_atoms)}")
        
        if len(h_atoms) == 0:
            print("⚠️  警告: 未找到H原子")
            return False
        
        # 检查H原子是否在末尾
        last_non_h_index = max(idx for idx, _ in non_h_atoms)
        first_h_index = min(idx for idx, _ in h_atoms)
        
        log_verbose(f"   最后非H原子位置: {last_non_h_index}", verbose)
        log_verbose(f"   第一个H原子位置: {first_h_index}", verbose)
        
        if first_h_index <= last_non_h_index:
            print("⚠️  警告: H原子未完全移动到末尾")
        
        # 创建重排序列
        log_verbose("🔄 开始重新排列...", verbose)
        
        # 分离不同类型的非H原子
        o_atoms = [(idx, row) for idx, row in non_h_atoms if str(row[0]).startswith('O(')]
        c_atoms = [(idx, row) for idx, row in non_h_atoms if str(row[0]).startswith('C(')]
        other_atoms = [(idx, row) for idx, row in non_h_atoms
                      if not str(row[0]).startswith('O(') and not str(row[0]).startswith('C(')]
        
        log_verbose(f"   O原子: {len(o_atoms)}, C原子: {len(c_atoms)}, 其他: {len(other_atoms)}", verbose)
        
        # 计算循环数量
        total_atoms = len(h_atoms) + len(non_h_atoms)
        # 总原子数 = 第一个(8) + 中间循环数×7 + 最后一个(8) + 剩余
        middle_cycles = (total_atoms - 16) // 7
        remaining_atoms = (total_atoms - 16) % 7
        
        log_verbose(f"   总原子数: {total_atoms}", verbose)
        log_verbose(f"   第一个循环: 8原子", verbose)
        log_verbose(f"   中间循环: {middle_cycles}个，每个7原子", verbose)
        log_verbose(f"   最后循环: 8原子", verbose)
        log_verbose(f"   剩余原子: {remaining_atoms}个", verbose)
        
        # 重排逻辑
        result_rows = []
        h_index = 0
        o_index = 0
        c_index = 0
        
        # 第一个循环: O H C H H C H H
        log_verbose("🔄 处理第一个循环: O H C H H C H H", verbose)
        for atom_type in FIRST_PATTERN:
            if atom_type == 'H' and h_index < len(h_atoms):
                result_rows.append(h_atoms[h_index][1])
                h_index += 1
            elif atom_type == 'O' and o_index < len(o_atoms):
                result_rows.append(o_atoms[o_index][1])
                o_index += 1
            elif atom_type == 'C' and c_index < len(c_atoms):
                result_rows.append(c_atoms[c_index][1])
                c_index += 1
        
        # 中间循环: O C H H C H H
        log_verbose(f"🔄 处理{middle_cycles}个中间循环: O C H H C H H", verbose)
        for cycle in range(middle_cycles):
            if verbose and cycle < 3:  # 只显示前3个中间循环的详细信息
                log_verbose(f"   处理第{cycle+1}个中间循环", verbose)
            for atom_type in MIDDLE_PATTERN:
                if atom_type == 'H' and h_index < len(h_atoms):
                    result_rows.append(h_atoms[h_index][1])
                    h_index += 1
                elif atom_type == 'O' and o_index < len(o_atoms):
                    result_rows.append(o_atoms[o_index][1])
                    o_index += 1
                elif atom_type == 'C' and c_index < len(c_atoms):
                    result_rows.append(c_atoms[c_index][1])
                    c_index += 1
        
        # 最后循环: O C H H C H H H
        log_verbose("🔄 处理最后循环: O C H H C H H H", verbose)
        for atom_type in LAST_PATTERN:
            if atom_type == 'H' and h_index < len(h_atoms):
                result_rows.append(h_atoms[h_index][1])
                h_index += 1
            elif atom_type == 'O' and o_index < len(o_atoms):
                result_rows.append(o_atoms[o_index][1])
                o_index += 1
            elif atom_type == 'C' and c_index < len(c_atoms):
                result_rows.append(c_atoms[c_index][1])
                c_index += 1
        
        # 添加剩余原子
        log_verbose(f"🔄 处理剩余原子...", verbose)
        while h_index < len(h_atoms):
            result_rows.append(h_atoms[h_index][1])
            h_index += 1
        while o_index < len(o_atoms):
            result_rows.append(o_atoms[o_index][1])
            o_index += 1
        while c_index < len(c_atoms):
            result_rows.append(c_atoms[c_index][1])
            c_index += 1
        for _, row in other_atoms:
            result_rows.append(row)
        
        # 创建结果DataFrame
        result_df = pd.DataFrame(result_rows)
        
        print(f"✅ 重排完成! 共处理 {len(result_df)} 行")
        log_verbose(f"   使用H原子: {h_index}/{len(h_atoms)}", verbose)
        log_verbose(f"   使用O原子: {o_index}/{len(o_atoms)}", verbose)
        log_verbose(f"   使用C原子: {c_index}/{len(c_atoms)}", verbose)
        
        if preview_only:
            print("\n📋 预览前24行 (前3个循环):")
            for i in range(min(24, len(result_df))):
                atom_name = result_df.iloc[i, 0]
                atom_type = atom_name.split('(')[0]
                print(f"{i+1:2}: {atom_type}", end=" ")
                
                # 标记循环边界
                if i == 7:
                    print(" <- 第一个循环结束")
                elif i == 14:
                    print(" <- 第一个中间循环结束")
                elif i == 21:
                    print(" <- 第二个中间循环结束")
                else:
                    print()
            
            print(f"\n预期模式:")
            print(f"第1-8位:   O H C H H C H H (第一个循环)")
            print(f"第9-15位:  O C H H C H H   (第一个中间循环)")
            print(f"第16-22位: O C H H C H H   (第二个中间循环)")
            return True
        
        # 保存文件
        log_verbose(f"💾 保存到 {output_file}...", verbose)
        result_df.to_excel(output_file, index=False, header=False)
        print(f"💾 文件已保存: {output_file}")
        
        return True
        
    except Exception as e:
        print(f"❌ 处理失败: {e}")
        return False

if __name__ == "__main__":
    input_file = sys.argv[1]
    output_file = sys.argv[2]
    verbose = sys.argv[3] == "True"
    preview_only = sys.argv[4] == "True"
    
    success = reorder_atoms(input_file, output_file, verbose, preview_only)
    sys.exit(0 if success else 1)
EOF
)

# 执行Python脚本
echo "🔄 开始处理..."
python3 -c "$PYTHON_SCRIPT" "$INPUT_FILE" "$OUTPUT_FILE" "$VERBOSE" "$PREVIEW_ONLY"

if [ $? -eq 0 ]; then
    # 计算处理时间
    END_TIME=$(date +%s)
    DURATION=$((END_TIME - START_TIME))
    
    echo ""
    echo "================================"
    echo "✅ 处理完成!"
    echo "⏱️  用时: ${DURATION}秒"
    
    if [ "$PREVIEW_ONLY" = false ]; then
        echo "📁 输出文件: $OUTPUT_FILE"
        
        # 显示文件信息
        if [ -f "$OUTPUT_FILE" ]; then
            FILE_SIZE=$(du -h "$OUTPUT_FILE" | cut -f1)
            echo "📦 文件大小: $FILE_SIZE"
        fi
    fi
    
else
    echo ""
    echo "❌ 处理失败"
    exit 1
fi
