import requests
import pandas as pd

# Замените API_KEY на ваш собственный
API_KEY = ''

def get_lists():
    url = f'https://api.unisender.com/ru/api/getLists?format=json&api_key={API_KEY}'
    response = requests.get(url)
    data = response.json()

    list_name_dict = {}
    for result in data['result']:
        list_name_dict[result['id']] = result['title']

    return list_name_dict


def get_campaigns(from_time, to_time):
    url = f'https://api.unisender.com/ru/api/getCampaigns?format=json&api_key={API_KEY}&from={from_time}&to={to_time}'
    response = requests.get(url)
    data = response.json()
    data_shortened = [{k: v for k, v in i.items() if k in ["id","sender_name", "subject", "sender_email", "list_id"]} for i in data['result']]
    return data_shortened


def get_campaign_stats(campaign, list_dict):
    url = f'https://api.unisender.com/ru/api/getCampaignCommonStats?format=json&api_key={API_KEY}&campaign_id={campaign["id"]}'
    response = requests.get(url)
    response.raise_for_status()

    data = response.json()
    data = {k: v for k, v in data['result'].items()}
    data['sender_name'] = campaign['sender_name']
    data['subject'] = campaign['subject']
    data['sender_email'] = campaign['sender_email']
    # add list title to the stats
    data['list_name'] = list_dict.get(campaign['list_id'], 'Произошла ошибка при получении имени списка')

    return data


def save_to_excel(data, filename):
    df = pd.DataFrame(data)
    df = df[['sender_name', 'subject', 'sender_email', 'list_name', 'total', 'sent', 'delivered', 'read_all', 'clicked_all', 'unsubscribed', 'spammed', 'read_unique', 'clicked_unique', 'spam']]
    df.to_excel(filename, index=False)

if __name__ == '__main__':
    #Введите диапазон дат отправки
    from_time = '2023-10-01 00:00:01'
    to_time = '2023-10-13 23:59:59'

    list_dict = get_lists()
    campaigns = get_campaigns(from_time, to_time)
    campaign_stats = []

    for campaign in campaigns:
        try:
            stats = get_campaign_stats(campaign, list_dict)
            campaign_stats.append(stats)
        except requests.HTTPError as err:
            print(f"Error occurred while fetching campaign {campaign['id']} statistics: {err}")

    filename = 'unisender_stats.xlsx'
    save_to_excel(campaign_stats, filename)
    print(f"Статистика сохранена в файле: {filename}")
