#!/usr/bin/env python
# coding: utf-8

# In[11]:


import requests
import pandas as pd

salvar_relatorio = 'C:\\Users\\francisco.sousa\\Documents\\Base_de_dados\\Relatórios_asana\\'

# Seu token de acesso pessoal do Asana
ASANA_TOKENS = ['2/1206489651519938/1206548070617399:4a042f79fc1a43680312be127c5e7214',
                '2/1207707413904280/1208169555336481:3b71a764594b7cc5dcff0cfc66c59116']
#     '2/1203537912621808/1206823689136115:855b02cd44d09db3d21edcd2a55364f9']

# ,
#                 ,
#                 '2/1203536821747205/1206784812211643:8baa6bc5806f86ea0adc627de314cd29',
#                 '2/1206489651519938/1206548070617399:34124f7d6cbc049dc3d6bb9676239a1f',
#                 '2/1203536254427118/1206718372686135:fc87208a2ee3c57a18d130a93b3f3464',
#                 '2/1201827646933243/1206659524655577:60ea598f18e3bd9f8d43a12ef8b507f9',
#                 '2/1203519038369657/1206726122669856:007f2d4910eb7fe9c390d2e02600737a']

# URL da API do Asana para listar tarefas
ASANA_API_URL = 'https://app.asana.com/api/1.0/tasks'

# Lista de parâmetros da solicitação (cada dicionário representa um conjunto de parâmetros)
request_params_list = [
    {
        'workspace': '1200730268694126',
        'assignee': 'me',
#         'opt_fields': "archived,assignee.name,projects.name,color,completed,completed_at,completed_by,completed_by.name,created_at,created_from_template,created_from_template.name,current_status,current_status.author,current_status.author.name,current_status.color,current_status.created_at,current_status.created_by,current_status.created_by.name,current_status.html_text,current_status.modified_at,current_status.text,current_status.title,current_status_update,current_status_update.resource_subtype,current_status_update.title,custom_field_settings,custom_field_settings.custom_field,custom_field_settings.custom_field.asana_created_field,custom_field_settings.custom_field.created_by,custom_field_settings.custom_field.created_by.name,custom_field_settings.custom_field.currency_code,custom_field_settings.custom_field.custom_label,custom_field_settings.custom_field.custom_label_position,custom_field_settings.custom_field.date_value,custom_field_settings.custom_field.date_value.date,custom_field_settings.custom_field.date_value.date_time,custom_field_settings.custom_field.description,custom_field_settings.custom_field.display_value,custom_field_settings.custom_field.enabled,custom_field_settings.custom_field.enum_options,custom_field_settings.custom_field.enum_options.color,custom_field_settings.custom_field.enum_options.enabled,custom_field_settings.custom_field.enum_options.name,custom_field_settings.custom_field.enum_value,custom_field_settings.custom_field.enum_value.color,custom_field_settings.custom_field.enum_value.enabled,custom_field_settings.custom_field.enum_value.name,custom_field_settings.custom_field.format,custom_field_settings.custom_field.has_notifications_enabled,custom_field_settings.custom_field.is_formula_field,custom_field_settings.custom_field.is_global_to_workspace,custom_field_settings.custom_field.is_value_read_only,custom_field_settings.custom_field.multi_enum_values,custom_field_settings.custom_field.multi_enum_values.color,custom_field_settings.custom_field.multi_enum_values.enabled,custom_field_settings.custom_field.multi_enum_values.name,custom_field_settings.custom_field.name,custom_field_settings.custom_field.number_value,custom_field_settings.custom_field.people_value,custom_field_settings.custom_field.people_value.name,custom_field_settings.custom_field.precision,custom_field_settings.custom_field.resource_subtype,custom_field_settings.custom_field.text_value,custom_field_settings.custom_field.type,custom_field_settings.is_important,custom_field_settings.parent,custom_field_settings.parent.name,custom_field_settings.project,custom_field_settings.project.name,custom_fields,custom_fields.date_value,custom_fields.date_value.date,custom_fields.date_value.date_time,custom_fields.display_value,custom_fields.enabled,custom_fields.enum_options,custom_fields.enum_options.color,custom_fields.enum_options.enabled,custom_fields.enum_options.name,custom_fields.enum_value,custom_fields.enum_value.color,custom_fields.enum_value.enabled,custom_fields.enum_value.name,custom_fields.is_formula_field,custom_fields.multi_enum_values,custom_fields.multi_enum_values.color,custom_fields.multi_enum_values.enabled,custom_fields.multi_enum_values.name,custom_fields.name,custom_fields.number_value,custom_fields.resource_subtype,custom_fields.text_value,custom_fields.type,default_access_level,default_view,due_date,due_on,followers,followers.name,html_notes,icon,members,members.name,minimum_access_level_for_customization,minimum_access_level_for_sharing,modified_at,name,notes,offset,owner,path,permalink_url,project_brief,public,start_on,team,team.name,uri,workspace,workspace.name", # list[str] | This endpoint returns a compact resource, which excludes some properties by default. To include those optional properties, set this query parameter to a comma-separated list of the properties you wish to include.
        'opt_fields': 'projects.name,workspace.name,id,name,assignee.name,completed,due_on,completed_at,modified_at,created_at,num_subtasks'
        
    },
    # Adicione mais conjuntos de parâmetros conforme necessário
]


data = []

# Iterar sobre tokens e parâmetros da solicitação
for token in ASANA_TOKENS:
    # Cabeçalhos para autenticação na API
    headers = {
        'Authorization': f'Bearer {token}'
    }

    # Iterar sobre conjuntos de parâmetros da solicitação
    for params in request_params_list:
        # Fazendo a solicitação GET para listar tarefas
        response = requests.get(ASANA_API_URL, headers=headers, params=params)

        # Verificando se a solicitação foi bem-sucedida
        if response.status_code == 200:
            tasks = response.json().get('data', [])
            for task in tasks:
                data.append({
                    'Token': token,
                    'nome_projeto': task['projects'][0]['name'] if task['projects'] else 'Sem projeto',
                    'nome_subtarefa': task['num_subtasks'],
                    'rede': task['workspace']['name'],
                    'Tarefa': task['name'],
                    'Responsável': task['assignee']['name'],
                    'Status da tarefa': task['completed'],
                    'Prazo estimado': task['due_on'],
                    'Concluída em': task['completed_at'],
                    'modified_at': task['modified_at'],
                    'created_at': task['created_at']                  
                })
        else:
            print(f"Erro ao acessar a API do Asana com o token {token}: {response.status_code} - {response.text}")

# Criar DataFrame final
tabela = pd.DataFrame(data)
tabela.to_excel(f'{salvar_relatorio}Relatório_tarefas_asana.xlsx', index=False)


# In[12]:


tabela


# In[ ]:




