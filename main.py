import requests
import pandas as pd

# Замените API_KEY на ваш собственный
API_KEY = ''

def get_campaigns(from_time, to_time):
    url = f'https://api.unisender.com/ru/api/getCampaigns?format=json&api_key={API_KEY}&from={from_time}&to={to_time}'
    response = requests.get(url)
    data = response.json()
    # Выбираем только нужную информацию
    data_shortened = [{k: v for k, v in i.items() if k in ["id","sender_name", "subject", "sender_email"]} for i in data['result']]
    return data_shortened


def get_campaign_stats(campaign):
   url = f'https://api.unisender.com/ru/api/getCampaignCommonStats?format=json&api_key={API_KEY}&campaign_id={campaign["id"]}'
   response = requests.get(url)

   # Raise an exception if status code indicates an error
   response.raise_for_status()

   data = response.json()
   # De-namespace the API's keys
   data = {k: v for k, v in data['result'].items()}

   # Add sender_name and sender_email to stats
   data['sender_name'] = campaign['sender_name']
   data['subject'] = campaign['subject']
   data['sender_email'] = campaign['sender_email']

   return data


def save_to_excel(data, filename):
    df = pd.DataFrame(data)
    # переупорядочение колонок
    df = df[['sender_name', 'subject', 'sender_email', 'total', 'sent', 'delivered', 'read_all', 'clicked_all', 'unsubscribed', 'spammed', 'read_unique', 'clicked_unique', 'spam']]
    df.to_excel(filename, index=False)

if __name__ == '__main__':
    # Введите даты в формате "YYYY-MM-DD HH:MM:SS"
    from_time = '2023-10-01 00:00:01'
    to_time = '2023-10-13 23:59:59'

    campaigns = get_campaigns(from_time, to_time)
    campaign_stats = []

    for campaign in campaigns:
        try:
            stats = get_campaign_stats(campaign)
            campaign_stats.append(stats)
        except requests.HTTPError as err:
            print(f"Error occurred while fetching campaign {campaign['id']} statistics: {err}")

    filename = 'unisender_stats.xlsx'
    save_to_excel(campaign_stats, filename)
    print(f"Статистика сохранена в файле: {filename}")

