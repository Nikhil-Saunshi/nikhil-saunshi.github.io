import json
import matplotlib.pyplot as plt
import pprint

company_id = '5809bc2777b61a00034ada25'
with open('chart_data.json') as f:
    data = json.load(f)
companies = data['chart'][0]['companies']


def get_graph_data(data, company_id):
    company_data = []
    companies = data['chart'][0]['companies']
    for company in companies:
        if company['_id'] == company_id:
            company_data = company['data']
            company_name = company['name']
            break
    else:
        print("Company ID not found")
        return
    # Filtered data from JSON to generate revenue vs date chart
    yearly_data = []
    quarterly_data = []
    for value_data in company_data:
        label = value_data['label']
        value = value_data['value']
        if '/' in label:
            quarterly_data.append((label, value))
        else:
            yearly_data.append((label, value))
    # pprint.pprint(company_data)
    print(yearly_data)
    print(quarterly_data)

    yearly_data = sorted(yearly_data, key=lambda x: int(x[0]))
    quarterly_data = sorted(quarterly_data, key=lambda x: int(x[0].split('/')[0]))
    print(yearly_data)
    print(quarterly_data)
    return yearly_data, quarterly_data, company_name


yearly_data, quarterly_data, company_name = get_graph_data(data, company_id)


def generate_graph(yearly_data, quarterly_data, company_name):
    fig, ax = plt.subplots(figsize=(5, 5))
    fig.subplots_adjust(bottom=0.15, left=0.2)
    ax.bar([val[0] for val in yearly_data], [val[1] for val in yearly_data], color=(0.2, 0.4, 0.6, 0.6))
    ax.set_title('Yearly revenue chart of ' + company_name)
    ax.set_xlabel('Year')
    ax.set_ylabel('Revenue Value in bn')
    fig.savefig('yearly_chart.jpg')

    fig, ax = plt.subplots(figsize=(5, 5))
    fig.subplots_adjust(bottom=0.15, left=0.2)
    ax.bar([val[0] for val in quarterly_data], [val[1] for val in quarterly_data], color=(0.2, 0.4, 0.6, 0.6))
    ax.set_title('Quarterly revenue chart of ' + company_name)
    ax.set_xlabel('Year')
    ax.set_ylabel('Revenue Value in bn')
    fig.savefig('2019_quarterly_chart.jpg')

    print("The graphs are ready to use now!!!!")


generate_graph(yearly_data, quarterly_data, company_name)
