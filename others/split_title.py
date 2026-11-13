import pandas as pd
import os

def split_combined_episodes(excel_path='video_detail.xlsx'):
    if not os.path.exists(excel_path):
        print(f"文件 {excel_path} 不存在")
        return

    # 讀取 Excel 文件
    df = pd.read_excel(excel_path, sheet_name='Episodes')

    new_rows = []

    for index, row in df.iterrows():
        season_number = row['Season Number']
        episode_number = row['Episode Number']
        episode_runtime = row['Episode Runtime']
        episode_title = row['Episode Title']
        episode_description = row['Episode Description']

        if '/' in episode_title or '／' in episode_title:
            titles = episode_title.replace('／', '/').split('/')
            for i, title in enumerate(titles):
                new_episode_number = episode_number + i
                new_rows.append([season_number, new_episode_number, episode_runtime, title.strip(), episode_description])
        else:
            new_rows.append([season_number, episode_number, episode_runtime, episode_title, episode_description])

    new_df = pd.DataFrame(new_rows, columns=['Season Number', 'Episode Number', 'Episode Runtime', 'Episode Title', 'Episode Description'])

    # 修正順延後的集數編號
    new_df = new_df.sort_values(by=['Season Number', 'Episode Number']).reset_index(drop=True)
    grouped = new_df.groupby('Season Number')
    
    final_rows = []
    for season, group in grouped:
        episode_number = 1
        for _, row in group.iterrows():
            final_rows.append([season, episode_number, row['Episode Runtime'], row['Episode Title'], row['Episode Description']])
            episode_number += 1

    final_df = pd.DataFrame(final_rows, columns=['Season Number', 'Episode Number', 'Episode Runtime', 'Episode Title', 'Episode Description'])

    # 寫入新的數據到 Excel 文件
    with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        final_df.to_excel(writer, sheet_name='Episodes', index=False)

    print(f"集數處理完成，更新後的數據已保存到 {excel_path}")

if __name__ == "__main__":
    split_combined_episodes()