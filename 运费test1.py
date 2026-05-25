import pandas as pd
import re
import os

file_path = '321.xlsx'
remark_file_path = '123_output.xlsx'  # 备注表文件路径

def extract_to_notepad_format():
    print("--- 正在进行智能物流费用提取 (防遗漏增强版) ---")

    try:
        # ==================================================
        # 读取备注表 (123_output) 并进行深度清洗
        # ==================================================
        remark_list = []  # 存储结构: [(原始备注, 纯净用来对比的备注)]
        if os.path.exists(remark_file_path):
            try:
                # 默认读取第一张工作表
                df_remarks = pd.read_excel(remark_file_path)
                if '备注' in df_remarks.columns:
                    raw_remarks = df_remarks['备注'].dropna().astype(str).unique().tolist()
                    for r in raw_remarks:
                        # 去掉所有空格、换行、制表符，只留纯文本用来做后期的安全比对
                        clean_r = re.sub(r'\s+', '', r)
                        remark_list.append((r, clean_r))
                    print(f"✅ 成功加载备注表，共获取 {len(remark_list)} 条有效备注基准")
                else:
                    print(f"❌ 错误：在【{remark_file_path}】中没有找到名为“备注”的列！请检查表头字样是否完全一致。")
            except Exception as e:
                print(f"⚠️ 读取备注表失败，错误原因：{e}")
        else:
            print(f"⚠️ 未找到【{remark_file_path}】文件，程序将无法追加备注。")

        # ==================================================
        # 读取主Excel
        # ==================================================
        df = pd.read_excel(file_path, header=None)
        output_blocks = []

        # ==================================================
        # 状态变量
        # ==================================================
        state = "SEARCH_HEADER"
        block_order = None
        block_city = None
        block_express_name = "韵达"

        # 更宽松的PO匹配
        order_pattern = re.compile(r'(PO-[A-Za-z0-9\-]+)')

        provinces = [
            '湖南', '湖北', '广东', '山东', '河南', '广西',
            '重庆', '四川', '江苏', '浙江', '安徽', '福建',
            '江西', '北京', '天津', '上海', '河北', '山西',
            '辽宁', '吉林', '黑龙江', '陕西', '甘肃', '青海',
            '贵州', '云南', '海南', '内蒙古', '西藏',
            '宁夏', '新疆'
        ]

        # ==================================================
        # 开始扫描
        # ==================================================
        for index, row in df.iterrows():
            # 当前行有效内容
            row_cells = [
                str(cell).strip()
                for cell in row
                if pd.notnull(cell) and str(cell).strip() != ''
            ]
            row_str = " ".join(row_cells)

            # ==================================================
            # 1. 检测仓开始（只认PO）
            # ==================================================
            order_match = order_pattern.search(row_str)

            if order_match:
                block_order = order_match.group(1)

                # 提取仓名前缀
                prefix_part = row_str.split(block_order)[0].strip()

                # 去掉省份
                for p in provinces:
                    if prefix_part.startswith(p):
                        prefix_part = prefix_part[len(p):].strip()
                        break

                # 清洗特殊字符
                prefix_part = re.sub(r'[总件：:\s]', '', prefix_part)

                # 取前两个字作为城市
                block_city = (
                    prefix_part[:2]
                    if len(prefix_part) >= 2
                    else prefix_part
                )

                # 防止空城市
                if not block_city:
                    block_city = "未知"

                # 重置
                block_express_name = "韵达"
                state = "SEARCH_COLUMNS"
                continue

            # ==================================================
            # 2. 查找快递名称
            # ==================================================
            if state in ["SEARCH_COLUMNS", "SEARCH_REAL_PAY"]:
                for cell in row_cells:
                    if "费用" in cell:
                        if "物流" in cell or "壹米" in cell:
                            continue

                        express_name = (
                            cell
                            .replace("费用", "")
                            .replace("快递", "")
                            .strip()
                        )
                        if express_name:
                            block_express_name = express_name

                state = "SEARCH_REAL_PAY"

            # ==================================================
            # 3. 查找实付
            # ==================================================
            if state == "SEARCH_REAL_PAY":
                if "实付" in row_str:
                    valid_numbers = []

                    for cell in row:
                        if pd.notnull(cell):
                            try:
                                val = float(str(cell).strip())
                                valid_numbers.append(val)
                            except:
                                pass

                    kuaidi_val = valid_numbers[0] if len(valid_numbers) >= 1 else 0
                    wuliu_val = valid_numbers[1] if len(valid_numbers) >= 2 else 0

                    kd_str = f"{round(kuaidi_val, 2):.2f}".rstrip('0').rstrip('.')
                    wl_str = f"{round(wuliu_val, 2):.2f}".rstrip('0').rstrip('.')

                    # ==================================================
                    # 强力防遗漏比对逻辑 (升级)
                    # ==================================================
                    matched_remark = ""
                    if block_city and block_city != "未知":
                        # 同样把提取出来的城市名去掉可能存在的杂质空格
                        clean_search_city = re.sub(r'\s+', '', block_city)
                        
                        for orig_r, clean_r in remark_list:
                            if clean_search_city in clean_r:
                                matched_remark = orig_r
                                break

                    # 生成文本
                    block_text = (
                        f"{block_order}\n"
                        f"辛苦下单邮费链接备注："
                        f"天猫美团{block_city}仓{block_express_name}\n"
                        f"快递{kd_str} 物流{wl_str}"
                    )
                    
                    if matched_remark:
                        block_text += f"\n{matched_remark}"
                        print(f"【成功提取】{block_order} -> 已成功配对 [{block_city}] 仓备注")
                    else:
                        block_text += f"\n⚠️ [暂无备注，请手动核对]"
                        print(f"❌【匹配失败】{block_order} -> 在备注表中未找到任何包含“{block_city}”的行！")

                    output_blocks.append(block_text)

                    # 重置状态
                    state = "SEARCH_HEADER"
                    block_order = None
                    block_city = None

        # ==================================================
        # 输出TXT
        # ==================================================
        if output_blocks:
            final_text = "\n\n".join(output_blocks)
            with open("物流费用提取结果.txt", "w", encoding="utf-8") as f:
                f.write(final_text)

            print("\n🎉【全部完成】")
            print(f"总计处理：{len(output_blocks)} 个仓")
            print("请检查上方控制台是否有 ❌ 提示，如有请核对数据源。")
        else:
            print("\n❌ 未提取到任何数据，请检查Excel格式")

    except Exception as e:
        print(f"\n运行出错：{e}")

if __name__ == '__main__':
    extract_to_notepad_format()
    input("\n按任意键退出程序...")
