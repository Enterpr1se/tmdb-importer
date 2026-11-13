import logging

def select_tw_hk_sheet(title, description):
    """
    根據標題和描述選擇 tw_hk_name.xlsx 的 sheet 名稱，忽略英文大小寫
    
    Args:
        title (str): 標題字符串
        description (str): 描述字符串
    
    Returns:
        str: 選擇的 sheet 名稱（例如 'Default', 'Gundam', 'One Piece', 'Pokemon'）
    """
    title_str = title if isinstance(title, str) else ''
    desc_str = description if isinstance(description, str) else ''
    
    # 將標題和描述轉為小寫以進行大小寫不敏感匹配
    title_lower = title_str.lower()
    desc_lower = desc_str.lower()
    
    # 定義關鍵詞映射
    sheet_keywords = {
        'Pokemon': ['pokemon', '寵物小精靈', '寶可夢', '神奇寶貝', '皮卡丘'],
        'Gundam': ['gundam', '鋼彈', '高達'],
        'One Piece': ['海賊', '航海王'],
        'Dragon Ball': ['Dragon Ball', '龍珠']
    }
    
    # 檢查標題和描述中的關鍵詞
    for sheet_name, keywords in sheet_keywords.items():
        for keyword in keywords:
            keyword_lower = keyword.lower()  # 確保關鍵詞也轉為小寫
            if keyword_lower in title_lower or keyword_lower in desc_lower:
                logging.info(f"檢測到關鍵詞 '{keyword}'，選擇 Sheet: {sheet_name}")
                return sheet_name
    
    # 如果未匹配任何關鍵詞，返回默認 sheet
    logging.info("未檢測到任何關鍵詞，選擇 Default Sheet")
    return 'Default'