import re

def parse_single_shorthand(text: str):
    """解析单个速记码"""
    text = text.strip()
    low = text.lower()
    
    print(f"DEBUG: 输入文本: '{text}', 小写: '{low}'")  # 调试输出
    
    # 删除 dd
    if low == 'dd':
        return None

    # 分离备注（首个汉字及以后）
    m = re.search(r'[\u4e00-\u9fff]', text)
    if m:
        idx = m.start()
        prefix, remark = text[:idx].strip(), text[idx:].strip()
    else:
        prefix, remark = text, ''

    print(f"DEBUG: 前缀: '{prefix}', 备注: '{remark}'")  # 调试输出

    # 单独数字不变
    if re.fullmatch(r'\d+', prefix):
        parsed = prefix
    else:
        parsed = parse_prefix(prefix.lower()) or prefix
        print(f"DEBUG: 解析结果: '{parsed}'")  # 调试输出

    # 附加备注（紧跟在处理后信息后面，没有空格）
    if remark:
        return f"{parsed}（{remark}）"
    return parsed

# 测试代码
test_cases = [
    "10kv@1111y",
    "220kv@2222y", 
    "10kv111y",
    "10kv@1",
    "dd"
]

for case in test_cases:
    result = parse_single_shorthand(case)
    print(f"输入: '{case}' -> 输出: '{result}'")