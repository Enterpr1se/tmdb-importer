import pandas as pd

def update_season_numbers(seasons_df):
    # 提取 Season Number 從 Season Name 中的 "第 x 季" 或 "第 x 輯"
    season_numbers = seasons_df['Season Name'].str.extract(r'第 (\d+) 季|第 (\d+) 輯').bfill(axis=1).iloc[:, 0]
    season_numbers = pd.to_numeric(season_numbers, errors='coerce')
    # 僅更新以 "第 x 季" 或 "第 x 輯" 開頭的 Season Name
    mask = ~season_numbers.isna()
    seasons_df.loc[mask, 'Season Number'] = season_numbers[mask].astype(int)
    return seasons_df

def update_episodes_numbers(episodes_df, seasons_df):
    # 創建 Season Name 到 Season Number 的映射
    season_mapping = dict(zip(seasons_df['Season Name'], seasons_df['Season Number']))
    # 創建舊的 Season Number 到新的 Season Number 的映射
    old_to_new_season_number = {index + 1: season_mapping.get(row['Season Name'], row['Season Number']) for index, row in seasons_df.iterrows()}
    # 更新 Episodes 中的 Season Number
    episodes_df['Season Number'] = episodes_df['Season Number'].map(old_to_new_season_number)
    return episodes_df

def process_excel(input_file, output_file):
    # 讀取輸入文件
    excel_file = pd.ExcelFile(input_file)
    titles_df = excel_file.parse('Title')
    seasons_df = excel_file.parse('Seasons')
    episodes_df = excel_file.parse('Episodes')

    # 更新 Season Number
    updated_seasons_df = update_season_numbers(seasons_df.copy())
    updated_episodes_df = update_episodes_numbers(episodes_df.copy(), updated_seasons_df)

    # 保存到輸出文件（覆蓋原文件）
    with pd.ExcelWriter(output_file, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        titles_df.to_excel(writer, sheet_name='Title', index=False)
        updated_seasons_df.to_excel(writer, sheet_name='Seasons', index=False)
        updated_episodes_df.to_excel(writer, sheet_name='Episodes', index=False)

# 使用示例
if __name__ == "__main__":
    input_file_path = './video_detail.xlsx'
    output_file_path = './video_detail.xlsx'
    process_excel(input_file_path, output_file_path)
