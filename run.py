import os
import sys

import openai
import tabula
import pandas as pd

INSTRUCTION = """I'm going to provide a list of items. Your task is to identify the item that matches most perfectly to the given string. Even if you're not confident, you must pick an item. Only return the index of the matching string in the list without any explanations. In fact, the returned output must be a single integer. Below are some examples.

Names: ['5`JOE-3`BHQ1', '5`CY5.5-3`BHQ2(0.2)', 'Modified primer synthesis 0.2 umoles', 'Modified 0.2 umoles oligo purification HPLC', '5`CY5 -3`BHQ-2']
String: '5`JOE-3`BHQ1'
Return: 0

Names: ['5`JOE-3`BHQ1', '5`CY5.5-3`BHQ2(0.2)', 'Modified primer synthesis 0.2 umoles', 'Modified 0.2 umoles oligo purification HPLC', '5`CY5 -3`BHQ-2']
String: '5`CY5-3`BHQ2'
Return: 4

Names: ['5`JOE-3`BHQ1', '5`CY5.5-3`BHQ2(0.2)', 'Modified primer synthesis 0.2 umoles', 'Modified 0.2 umoles oligo purification HPLC', '5`CY5 -3`BHQ-2']
String: '5`CY5.5-3`BHQ2'
Return: 1

Names: ['5`CY5.5-3`BHQ2(0.2)', 'Modified primer synthesis 0.2 umoles', 'Modified 0.2 umoles oligo purification HPLC']
String: '5`CY5.5-3`BHQ2'
Return: 0

Names: ['Modified primer synthesis 0.2 umoles', 'Modified 0.2 umoles oligo purification HPLC', '5`FAM-3`BHQ1', '5`CY5 -3`BHQ-2']
String: '5`CY5-3`BHQ2'
Return: 3

Now it's your turn:"""

def compute_cost_syn(df, amount):
    s = df[df.품명 == f'Modified primer synthesis {amount} umoles']['단가']
    if len(s) != 1:
        raise ValueError("Multiple rows found for primer synthesis")
    return s.values[0]

def compute_cost_mod(df, mod5, mod3):
    """
    Uses OpenAI API to find the best matching string in the list of names.
    """
    prompt = INSTRUCTION + f"\n\nNames: {df.품명.to_list()}\nString: '{mod5}-{mod3}'"
    bot_response = openai.ChatCompletion.create(
        model='gpt-3.5-turbo',
        # model='gpt-4',
        messages=[{"role": "system", "content": prompt}]
    )
    i = bot_response["choices"][0]["message"]["content"]
    # print('-' * 80)
    # print(prompt)
    # print(i)
    return df.iloc[int(i)]['단가']

def compute_cost_pur(df, amount):
    s = df[df.품명 == f'Modified {amount} umoles oligo purification HPLC']['단가']
    if len(s) != 1:
        raise ValueError("Multiple rows found for oligo purification")
    return s.values[0]

if __name__ == '__main__':
    openai.api_key = os.environ['OPENAI_API_KEY']

    order_dir = sys.argv[1]
    order_number = os.path.basename(order_dir)

    for root, dirs, files in os.walk(order_dir):
        for file in files:
            if "고객주문내역" in file:
                excel_file = file # 주문조회
            elif "오더시스템" in file:
                pdf_file = file # 출하조회

    df1 = pd.read_excel(f'{order_dir}/{excel_file}', engine='openpyxl')

    tables = tabula.read_pdf(f'{order_dir}/{pdf_file}', pages='all') 
    df2 = tables[1]
    df2 = df2.drop(df2.columns[[1, 2]], axis=1)
    df2 = df2.drop(df2.tail(2).index)
    df2.columns.values[0] = '품명'
    df2.단가 = df2.단가.str.replace(',', '').astype(int)
    df2.공급가액 = df2.공급가액.str.replace(',', '').astype(int)
    df2.세액 = df2.세액.str.replace(',', '').astype(int)

    actual_total = df2.공급가액.sum() + df2.세액.sum()
    expected_total = 0

    data = {}

    for i, row in df1.iterrows():
        oligo = row['Oligo Name']
        amount = row['Amount']
        if amount == 1.0:
            amount = 1
        elif amount == 0.2:
            amount = 0.2
        else:
            raise ValueError('Invalid amount: {amount}')
        mer = row['mer']
        mod5 = row["5`Mod"].strip()
        mod3 = row["3`Mod"].strip()
        cost_syn = compute_cost_syn(df2, amount)
        cost_mod = compute_cost_mod(df2, mod5, mod3)
        cost_pur = compute_cost_pur(df2, amount)
        price = cost_syn * mer + cost_mod + cost_pur
        tax = price / 10
        expected_total += price + tax
        data[oligo] = {
            '규격': f"{amount}umole",
            '수량': 1,
            '단가': '{:,}'.format(price),
            '공급가액': '{:,}'.format(price),
            '세액': '{:,}'.format(tax),
            '비고': os.path.basename(order_dir),
            '합성 단가': cost_syn,
            'mer 수': mer,
            '합성비': cost_syn * mer,
            'Modification': cost_mod,
            '정제': cost_pur,
        }

    if expected_total == actual_total:
        print(f"수주번호 [{order_number}] 금액이 정상적으로 처리되었습니다.")
    else:
        print(f"수주번호 [{order_number}] 금액에 오차가 발생하였습니다. 확인이 필요합니다.")

    df3 = pd.DataFrame(data).T
    df3 = df3.reset_index(names='품명')
    df3.to_excel(f'{order_dir}/거래명세서-{order_number}.xlsx', index=False)