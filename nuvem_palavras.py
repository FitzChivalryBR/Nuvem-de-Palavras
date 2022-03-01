import matplotlib.pyplot as plt
from wordcloud import WordCloud
import numpy as np
from PIL import Image
import nltk
from nltk.corpus import stopwords
import pandas as pd 
import time, os
from pathlib import Path

# pip install emojis matplotlib wordcloud

def tamanho_char_subject(subj):
    texto = subj['subject']
    size_txt = len(texto.replace(' ', ''))

    return size_txt

def tamanho_word_subject(subj):
    texto = subj['subject']
    size_txt = len(texto.split(' '))

    return size_txt

def ajustes_subj(subj):
    p = (subj['subject']).replace('[Test]:', '').lower()

    return p


def format_number(n):
    nx = float(str(n).replace(',', '.'))

    return nx

# TESTAR PALAVRAS JUNTAS!

nltk.download('stopwords')
sw_pt = stopwords.words('portuguese')
sw_eng = stopwords.words('english')

n_sw_pt = ['[Test]', 'Test', 'é', 'o', '0', '[Test]:', 'pra']
pontuais = ['!', "'", "[", "]", ':', '.', ',']
permitir = ['?']
sw_pt.extend(n_sw_pt)


for p in permitir:
    try:
        sw_pt.remove(p)
    except:
        pass

subjects = pd.read_csv('word_cloud.csv', encoding='UTF-8', sep=';')
subjects['tx_engj'] = subjects['tx_engj'].apply(format_number)
subjects['qt_char_subject'] = subjects.apply(tamanho_char_subject, axis=1)
subjects['qt_plvrs_subject'] = subjects.apply(tamanho_word_subject, axis=1) 
subjects['subject'] = subjects.apply(ajustes_subj, axis=1)



plvrs_subj = pd.DataFrame()
analise = list(subjects['ano'].unique())

print('\n\nAnalisando subjects por ano... Aguarde!')
Path('por_ano').mkdir(parents=True, exist_ok=True)

for filtro in analise:
    print(f'Processando Ano.: {filtro}')
    print('| [...] |')

    subjects_temp = subjects[subjects['ano'] == filtro]
    if subjects_temp.empty:
        continue
    #Métricas Gerais
    #Métricas Gerais
    #Métricas Gerais
    media_caract = subjects[subjects['ano'] == filtro]['qt_char_subject'].mean()
    media_palavras = subjects[subjects['ano'] == filtro]['qt_plvrs_subject'].mean()

    abertura = subjects[subjects['ano'] == filtro]['abertura'].sum()
    enviado = subjects[subjects['ano'] == filtro]['enviados'].sum()
    clique = subjects[subjects['ano'] == filtro]['cliques'].sum()
    optout = subjects[subjects['ano'] == filtro]['optout'].sum()

    tx_abertura = (abertura/enviado)*100
    tx_clique = (clique/abertura)*100
    tx_optout = (optout/abertura)*100


    
    p = []
    for li in range(len(subjects_temp)):
        df_subj = subjects_temp.iloc[li]

        p = (df_subj.subject).replace('[Test]:', '').lower()
        p = (df_subj.subject).replace(':', ' ').lower()

        for v in pontuais:
            p = p.replace(v, '').strip()


        p = p.split(' ')

        for i in range(len(p)):
            if str(p[i]) == '??':
                pass
            elif str(p[i]).endswith('?'):
                p[i] = p[i].replace('?', '')
            elif str(p[i]).endswith(','):
                p[i] = p[i].replace(',', '')
            elif str(p[i]).endswith('!'):
                p[i] = p[i].replace('!', '')
                
        for v in p:
            if v in sw_pt:
                while v in p:
                    p.remove(v)

        for v in p:
            if v in sw_pt:
                while v in p:
                    p.remove(v)

        for k in p:
            if k.strip() in sw_pt:
                continue
            else:
                valor = {'palavras_from_subj': k.strip().lower(),
                        'tx_engj': float(str(df_subj.tx_engj).replace(',', '.'))
                        #'p_abertura': float(str(df_subj.p_abertura).replace(',', '.'))
                        }
            
                plvrs_subj = plvrs_subj.append(valor, ignore_index=True, sort=False)
                # Verificar se há necessidade de remover palavras usadas apenas uma vez!


    # Taxa de Engajamento
    palavras_unicas = len(list(plvrs_subj['palavras_from_subj'].unique()))
    plvrs_subj['QTD_REPET'] = plvrs_subj.groupby(by=['palavras_from_subj']).transform('count')
    plvrs_subj['VALOR_PONDERADO'] = (plvrs_subj['tx_engj']*plvrs_subj['QTD_REPET'])/plvrs_subj['QTD_REPET'].values.sum()
    plvrs_subj.to_excel(f'por_ano\\total_palavras_tx_engj_{filtro}.xlsx', index=False)


    # (palavrax * qtd_repet_palavrax) / total_repet (3000)


    # Taxa de Abertura
    #plvrs_subj['QTD_REPET'] = plvrs_subj.groupby(by=['palavras_from_subj']).transform('count')
    #plvrs_subj['VALOR_PONDERADO'] = (plvrs_subj['p_abertura']*plvrs_subj['QTD_REPET'])/plvrs_subj['QTD_REPET'].values.sum()
    #plvrs_subj.to_excel('total_palavras_tx_abert.xlsx', index=False)



    # group by - wc
    df_media_palavras = plvrs_subj[plvrs_subj['palavras_from_subj'] != ''].groupby(by=['palavras_from_subj']).mean()
    dict_media_plvrs = df_media_palavras.to_dict()['VALOR_PONDERADO']
    plvrs_subj = pd.DataFrame()
    
    df_media_palavras.to_excel(f'por_ano\\wc_{filtro}.xlsx', index=True)


    # Estilo
    shape_gb = np.array(Image.open(r'shape_gb.png'))
    font_pers = r'Symbola.otf'

    wc = WordCloud(stopwords = sw_pt,
                mask = shape_gb,
                font_path=font_pers,
                background_color = "white",
                max_words = 3000,
                max_font_size = 500,
                random_state = 42,
                width = shape_gb.shape[1],
                height = shape_gb.shape[0],
                colormap='copper'
                #contour_color='black',
                #contour_width=12
                )


    # Gerando a nuvem de palavras
    # wc.generate(river)
    print('Gerando nuvem de palavras...')
    wc.generate_from_frequencies(dict_media_plvrs)
    plt.imshow(wc, interpolation="bilinear")
    plt.axis('off')
    print('Salvandor figura, aguarde...')
    plt.savefig(f"por_ano\\word_cloud_tx_engj_{filtro}.png", format="png", dpi=500)
    # plt.savefig("word_cloud_tx_abert.png", format="png", dpi=500)
    print('Finalizado!\n')

    valores = {
        'objeto_analise': filtro,
        'tx_abertura': tx_abertura,
        'tx_clique': tx_clique,
        'tx_optout': tx_optout,
        'media_palavras': media_palavras,
        'media_caract': media_caract,
        'palavras_unicas': palavras_unicas,
        'top20_palavras': list(wc.words_.keys())[:20],
        'bottom20_palavras': list(wc.words_.keys())[-20:]
    }

    df_export = pd.DataFrame.from_dict(valores)
    df_export.to_excel(f'por_ano\\dados_gerais_{filtro}.xlsx', index=False)
    wc = None
    df_export = None
    plvrs_subj = pd.DataFrame()



















# Categoria
# Categoria
# Categoria
# Categoria

plvrs_subj = pd.DataFrame()
analise = list(subjects['ano'].unique())
sub_analis = list(subjects['promo_categoria'].unique())


print('\n\nAnalisando subjects por Ano & Categoria... Aguarde!')
Path('por_categoria').mkdir(parents=True, exist_ok=True)


for filtro in analise:
    for cat in sub_analis:
        print(f'Processando Ano.: {filtro} & Categoria: {cat}')
        Path(f'por_categoria\\{filtro}\\{cat}').mkdir(parents=True, exist_ok=True)
        print('| [...] |')

        subjects_temp = subjects[(subjects['ano'] == filtro) & (subjects['promo_categoria'] == cat)]
        if subjects_temp.empty:
            continue

        #Métricas Gerais
        #Métricas Gerais
        #Métricas Gerais
        media_caract = subjects[(subjects['ano'] == filtro) & (subjects['promo_categoria'] == cat)]['qt_char_subject'].mean()
        media_palavras = subjects[(subjects['ano'] == filtro) & (subjects['promo_categoria'] == cat)]['qt_plvrs_subject'].mean()

        abertura = subjects[(subjects['ano'] == filtro) & (subjects['promo_categoria'] == cat)]['abertura'].sum()
        enviado = subjects[(subjects['ano'] == filtro) & (subjects['promo_categoria'] == cat)]['enviados'].sum()
        clique = subjects[(subjects['ano'] == filtro) & (subjects['promo_categoria'] == cat)]['cliques'].sum()
        optout = subjects[(subjects['ano'] == filtro) & (subjects['promo_categoria'] == cat)]['optout'].sum()

        tx_abertura = (abertura/enviado)*100
        tx_clique = (clique/abertura)*100
        tx_optout = (optout/abertura)*100

        p = []
        for li in range(len(subjects_temp)):
            df_subj = subjects_temp.iloc[li]

            p = (df_subj.subject).replace('[Test]:', '').lower()
            p = (df_subj.subject).replace(':', ' ').lower()

            for v in pontuais:
                p = p.replace(v, '').strip()


            p = p.split(' ')

            for i in range(len(p)):
                if str(p[i]) == '??':
                    pass
                elif str(p[i]).endswith('?'):
                    p[i] = p[i].replace('?', '')
                elif str(p[i]).endswith(','):
                    p[i] = p[i].replace(',', '')
                elif str(p[i]).endswith('!'):
                    p[i] = p[i].replace('!', '')
                    
            for v in p:
                if v in sw_pt:
                    while v in p:
                        p.remove(v)

            for v in p:
                if v in sw_pt:
                    while v in p:
                        p.remove(v)

            for k in p:
                if k.strip() in sw_pt:
                    continue
                else:
                    valor = {'palavras_from_subj': k.strip().lower(),
                            'tx_engj': float(str(df_subj.tx_engj).replace(',', '.'))
                            #'p_abertura': float(str(df_subj.p_abertura).replace(',', '.'))
                            }
                
                    plvrs_subj = plvrs_subj.append(valor, ignore_index=True, sort=False)


        # Taxa de Engajamento
        palavras_unicas = len(list(plvrs_subj['palavras_from_subj'].unique()))
        plvrs_subj['QTD_REPET'] = plvrs_subj.groupby(by=['palavras_from_subj']).transform('count')
        plvrs_subj['VALOR_PONDERADO'] = (plvrs_subj['tx_engj']*plvrs_subj['QTD_REPET'])/plvrs_subj['QTD_REPET'].values.sum()
        plvrs_subj.to_excel(f'por_categoria\\{filtro}\\{cat}\\total_palavras_tx_engj_{filtro}{cat}.xlsx', index=False)


        # group by - wc
        df_media_palavras = plvrs_subj[plvrs_subj['palavras_from_subj'] != ''].groupby(by=['palavras_from_subj']).mean()
        dict_media_plvrs = df_media_palavras.to_dict()['VALOR_PONDERADO']
        plvrs_subj = pd.DataFrame()

        df_media_palavras.to_excel(f'por_categoria\\{filtro}\\{cat}\\wc_{filtro}{cat}.xlsx', index=True)


        # Estilo
        shape_gb = np.array(Image.open(r'shape_gb.png'))
        font_pers = r'Symbola.otf'

        wc = WordCloud(stopwords = sw_pt,
                    mask = shape_gb,
                    font_path=font_pers,
                    background_color = "white",
                    max_words = 3000,
                    max_font_size = 500,
                    random_state = 42,
                    width = shape_gb.shape[1],
                    height = shape_gb.shape[0],
                    colormap='copper'
                    #contour_color='black',
                    #contour_width=12
                    )


        # Gerando a nuvem de palavras
        # wc.generate(river)
        print('Gerando nuvem de palavras...')
        wc.generate_from_frequencies(dict_media_plvrs)
        plt.imshow(wc, interpolation="bilinear")
        plt.axis('off')
        print('Salvandor figura, aguarde...')
        plt.savefig(f"por_categoria\\{filtro}\\{cat}\\word_cloud_tx_engj_{filtro}{cat}.png", format="png", dpi=500)
        # plt.savefig("word_cloud_tx_abert.png", format="png", dpi=500)
        print('Finalizado!\n')

        valores = {
            'objeto_analise': filtro,
            'tx_abertura': tx_abertura,
            'tx_clique': tx_clique,
            'tx_optout': tx_optout,
            'media_palavras': media_palavras,
            'media_caract': media_caract,
            'palavras_unicas': palavras_unicas,
            'top20_palavras': list(wc.words_.keys())[:20],
            'bottom20_palavras': list(wc.words_.keys())[-20:]
        }

        df_export = pd.DataFrame.from_dict(valores)
        df_export.to_excel(f'por_categoria\\{filtro}\\{cat}\\dados_gerais_{filtro}{cat}.xlsx', index=False)
        wc = None
        df_export = None
        plvrs_subj = pd.DataFrame()



















# Ação
# Ação
# Ação
# Ação

plvrs_subj = pd.DataFrame()
analise = list(subjects['ano'].unique())
sub_analis = list(subjects['tipo_acao'].unique())


print('\n\nAnalisando subjects por Ano & Ação... Aguarde!')
Path('por_acao').mkdir(parents=True, exist_ok=True)


for filtro in analise:
    for cat in sub_analis:
        
        print(f'Processando Ano.: {filtro} & Ação: {cat}')
        Path(f'por_acao\\{filtro}\\{cat}').mkdir(parents=True, exist_ok=True)
        print('| [...] |')

        subjects_temp = subjects[(subjects['ano'] == filtro) & (subjects['tipo_acao'] == cat)]
        if subjects_temp.empty:
            continue

        #Métricas Gerais
        #Métricas Gerais
        #Métricas Gerais
        media_caract = subjects[(subjects['ano'] == filtro) & (subjects['tipo_acao'] == cat)]['qt_char_subject'].mean()
        media_palavras = subjects[(subjects['ano'] == filtro) & (subjects['tipo_acao'] == cat)]['qt_plvrs_subject'].mean()

        abertura = subjects[(subjects['ano'] == filtro) & (subjects['tipo_acao'] == cat)]['abertura'].sum()
        enviado = subjects[(subjects['ano'] == filtro) & (subjects['tipo_acao'] == cat)]['enviados'].sum()
        clique = subjects[(subjects['ano'] == filtro) & (subjects['tipo_acao'] == cat)]['cliques'].sum()
        optout = subjects[(subjects['ano'] == filtro) & (subjects['tipo_acao'] == cat)]['optout'].sum()

        tx_abertura = (abertura/enviado)*100
        tx_clique = (clique/abertura)*100
        tx_optout = (optout/abertura)*100

        p = []
        for li in range(len(subjects_temp)):
            df_subj = subjects_temp.iloc[li]

            p = (df_subj.subject).replace('[Test]:', '').lower()
            p = (df_subj.subject).replace(':', ' ').lower()

            for v in pontuais:
                p = p.replace(v, '').strip()


            p = p.split(' ')

            for i in range(len(p)):
                if str(p[i]) == '??':
                    pass
                elif str(p[i]).endswith('?'):
                    p[i] = p[i].replace('?', '')
                elif str(p[i]).endswith(','):
                    p[i] = p[i].replace(',', '')
                elif str(p[i]).endswith('!'):
                    p[i] = p[i].replace('!', '')
                    
            for v in p:
                if v in sw_pt:
                    while v in p:
                        p.remove(v)

            for v in p:
                if v in sw_pt:
                    while v in p:
                        p.remove(v)

            for k in p:
                if k.strip() in sw_pt:
                    continue
                else:
                    valor = {'palavras_from_subj': k.strip().lower(),
                            'tx_engj': float(str(df_subj.tx_engj).replace(',', '.'))
                            #'p_abertura': float(str(df_subj.p_abertura).replace(',', '.'))
                            }
                
                    plvrs_subj = plvrs_subj.append(valor, ignore_index=True, sort=False)


        # Taxa de Engajamento
        palavras_unicas = len(list(plvrs_subj['palavras_from_subj'].unique()))
        plvrs_subj['QTD_REPET'] = plvrs_subj.groupby(by=['palavras_from_subj']).transform('count')
        plvrs_subj['VALOR_PONDERADO'] = (plvrs_subj['tx_engj']*plvrs_subj['QTD_REPET'])/plvrs_subj['QTD_REPET'].values.sum()

        cat = cat.replace('/SubFoco', '')
        plvrs_subj.to_excel(f'por_acao\\{filtro}\\{cat}\\total_palavras_tx_engj_{filtro}{cat}.xlsx', index=False)


        # group by - wc
        df_media_palavras = plvrs_subj[plvrs_subj['palavras_from_subj'] != ''].groupby(by=['palavras_from_subj']).mean()
        dict_media_plvrs = df_media_palavras.to_dict()['VALOR_PONDERADO']
        plvrs_subj = pd.DataFrame()

        df_media_palavras.to_excel(f'por_acao\\{filtro}\\{cat}\\wc_{filtro}{cat}.xlsx', index=True)


        # Estilo
        shape_gb = np.array(Image.open(r'shape_gb.png'))
        font_pers = r'Symbola.otf'

        wc = WordCloud(stopwords = sw_pt,
                    mask = shape_gb,
                    font_path=font_pers,
                    background_color = "white",
                    max_words = 3000,
                    max_font_size = 500,
                    random_state = 42,
                    width = shape_gb.shape[1],
                    height = shape_gb.shape[0],
                    colormap='copper'
                    #contour_color='black',
                    #contour_width=12
                    )


        # Gerando a nuvem de palavras
        # wc.generate(river)
        print('Gerando nuvem de palavras...')
        wc.generate_from_frequencies(dict_media_plvrs)
        plt.imshow(wc, interpolation="bilinear")
        plt.axis('off')
        print('Salvandor figura, aguarde...')
        plt.savefig(f"por_acao\\{filtro}\\{cat}\\word_cloud_tx_engj_{filtro}{cat}.png", format="png", dpi=500)
        # plt.savefig("word_cloud_tx_abert.png", format="png", dpi=500)
        print('Finalizado!\n')

        valores = {
            'objeto_analise': filtro,
            'tx_abertura': tx_abertura,
            'tx_clique': tx_clique,
            'tx_optout': tx_optout,
            'media_palavras': media_palavras,
            'media_caract': media_caract,
            'palavras_unicas': palavras_unicas,
            'top20_palavras': list(wc.words_.keys())[:20],
            'bottom20_palavras': list(wc.words_.keys())[-20:]
        }

        df_export = pd.DataFrame.from_dict(valores)
        df_export.to_excel(f'por_acao\\{filtro}\\{cat}\\dados_gerais_{filtro}{cat}.xlsx', index=False)
        wc = None
        df_export = None
        plvrs_subj = pd.DataFrame()



        