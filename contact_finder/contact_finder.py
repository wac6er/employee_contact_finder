import xlwings as xw



def not_not_main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    import pandas as pd
    import requests
    import sys

    from bs4 import BeautifulSoup
    import pandas as pd
    from io import StringIO
    import numpy as np
    import unidecode
    from time import sleep
    import time
    import bs4

    import smtplib
    import dns
    import dns.resolver
    import pandas as pd
    import socket
    import pprint


    def contact_info_finder(company):
        summary = {'name':[],'title':[]}
        email_check = {'a_first_name':[],'last_name':[],'address':[]}
        title = a

        summ = []

        url = 'https://www.google.com/search?q='+company+'+'+title+'+linkedin'
        html = requests.get(url).text
        soup = bs4.BeautifulSoup(html, "html.parser")
        for tag in soup.findAll("div", {"class": "egMi0 kCrYT"}):
                summ.append(tag.findNext().text)

        for xy in range(3):
            z = summ[xy]
            end = z.find('-')
            substring0 = z[:end]

            s = summ[xy]
            start = s.find('-')
            end = s.find('www.linkedin.com')
            substring = s[start:end]



            if len(substring0)>4:
                summary['name'].append(substring0)
                summary['title'].append(substring)
        return summary


    xx = {'titles':['chief of sales','vp of sales','ecommerce manager','strategic partnerships','chief marketing officer','national account manager','account manager','sales manager','regional account manager','affiliate program manager','business development manager','eccommerce marketing manager']}
    x = {'titles':['chief of sales','vp of sales','ecommerce manager','strategic partnerships','chief marketing officer','national account manager','account manager','sales manager','regional account manager','affiliate program manager','business development manager','eccommerce marketing manager']}
    x = pd.DataFrame(x)
    u = {'a':[],'title':[]}
    for i in range(len(x['titles'])):
        a = x['titles'][i]
        u['a'].append(contact_info_finder(wb.sheets[0].range('A2').value))

    aaaa = u['a']

    new_df = pd.DataFrame(u['a'][0])
    if len(new_df) ==1:
        tit1 = xx['titles'][0]
    else:
        print()
    new_df1 = pd.DataFrame(u['a'][1])
    if len(new_df1) ==1:
        tit2 = xx['titles'][1]
    else:
        print()
    new_df2 = pd.DataFrame(u['a'][2])
    if len(new_df2) ==1:
        tit3 = xx['titles'][2]
    else:
        print()
    new_df3 = pd.DataFrame(u['a'][3])
    if len(new_df3) ==1:
        tit4 = xx['titles'][3]
    else:
        print()
    new_df4 = pd.DataFrame(u['a'][4])
    if len(new_df4) ==1:
        tit5 = xx['titles'][4]
    else:
        print()
    new_df5 = pd.DataFrame(u['a'][5])
    if len(new_df5) ==1:
        tit6 = xx['titles'][5]
    else:
        print()

    peices = (new_df,new_df1,new_df2,new_df3,new_df4,new_df5)
    df = pd.concat(peices, ignore_index = True)




    wb.sheets[0].range('B1').value = df

def emails():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    import pandas as pd
    import requests
    import sys

    from bs4 import BeautifulSoup
    import pandas as pd
    from io import StringIO
    import numpy as np
    import unidecode
    from time import sleep
    import time
    import bs4

    import smtplib
    import dns
    import dns.resolver
    import pandas as pd
    import socket
    import pprint

    contact_names = pd.read_excel('/Users/willcarrington/demo/people_finder.xlsm')
    contact_names = contact_names['name'].str.split()
    first = []
    last = []
    adress = []
    for x in range(len(contact_names)):
        first.append(contact_names[x][0])
        last.append(contact_names[x][1])
        adress.append(wb.sheets[0].range('C22').value)
    email_check = {'first':[],'last':[],'address':[]}
    email_check['first'].append(first[0])
    email_check['last'].append(last[0])
    email_check['address'].append(adress[0])
    email_check = pd.DataFrame(email_check)

    First_name=email_check['first'][0]
    Last_name = email_check['last'][0]
    Company = email_check['address'][0]



    def email_checker(scale):
        email = {'check':[],'email':[]}
        for x in scale:
            email_address= x
            addressToVerify = email_address

            domain_name = email_address.split('@')[1]

            records = dns.resolver.query(domain_name, 'MX')
            mxRecord = records[0].exchange
            mxRecord = str(mxRecord)

            host = socket.gethostname()

            server = smtplib.SMTP()
        
            server.set_debuglevel(0)
            server.connect(mxRecord)
            server.helo(host)
            server.mail('me@domain.com')
            code, message = server.rcpt(str(addressToVerify))
            server.quit()

            if code == 250:
                email['check'].append('Y')
                email['email'].append(email_address)

        email = pd.DataFrame(email)
        return (email.head())

    def name1(s,g,Company):

        l = s.split()
        g = g.split()
        new = ""
        second = "" 
        third = ""
        fourth = ""
        fifth = ""
        sixth = ""
        seventh = ""
  
        for i in range(len(l)):
            s = l[i]
            g = g[i]
            new += (s[0:].lower())
            second += (s[0:].lower()+g[0:].lower())
            third += (s[0:1].lower()+g[0:].lower())
            fourth += (s[0:].lower()+g[0:1].lower())
            fifth += (g[0:1].lower()+s[0:].lower())
            sixth += (g[0:].lower()+s[0:1].lower())
            seventh += (g[0:].lower())
        
          
      
        return (new+'@'+Company, second+'@'+Company, third+'@'+Company, fourth+'@'+Company, fifth+'@'+Company, sixth+'@'+Company, seventh+'@'+Company+com)

    def name2(s,g,Company):

        l = s.split()
        g = g.split()
        new = ""
        second = "" 
        third = ""
        fourth = ""
        fifth = ""
        sixth = ""
        seventh = ""
        a= ""
        b= ""
        c= ""
        d= ""
        e= ""
  
        for i in range(len(l)):
            s = l[i]
            g = g[i]
            new += (s[0:].lower())
            second += (s[0:].lower()+'.'+g[0:].lower())
            third += (s[0:1].lower()+'.'+g[0:].lower())
            fourth += (s[0:].lower()+'.'+g[0:1].lower())
            fifth += (g[0:1].lower()+'.'+s[0:].lower())
            sixth += (g[0:].lower()+'.'+s[0:1].lower())
            seventh += (g[0:].lower())
            a += (s[0:].lower()+g[0:].lower())
            b += (s[0:1].lower()+g[0:].lower())
            c += (s[0:].lower()+g[0:1].lower())
            d += (g[0:1].lower()+s[0:].lower())
            e += (g[0:].lower()+s[0:1].lower())
          
      
        return (new+'@'+Company, second+'@'+Company, third+'@'+Company, fourth+'@'+Company, fifth+'@'+Company, sixth+'@'+Company, seventh+'@'+Company, a+'@'+Company, b+'@'+Company, c+'@'+Company, d+'@'+Company, e+'@'+Company+com)

    def name3(s,g,company):

        l = s.split()
        g = g.split()
        new = ""
        second = "" 
        third = ""
        fourth = ""
        fifth = ""
        sixth = ""
        seventh = ""
        a= ""
        b= ""
        c= ""
        d= ""
        e= ""
        a1= ""
        a2= ""
        a3= ""
        a4= ""
        a5= ""
        a11= ""
        a22= ""
        a33= ""
        a44= ""
        a55= ""
  
        for i in range(len(l)):
            s = l[i]
            g = g[i]
            new += (s[0:].lower())
            second += (s[0:].lower()+'.'+g[0:].lower())
            third += (s[0:1].lower()+'.'+g[0:].lower())
            fourth += (s[0:].lower()+'.'+g[0:1].lower())
            fifth += (g[0:1].lower()+'.'+s[0:].lower())
            sixth += (g[0:].lower()+'.'+s[0:1].lower())
            seventh += (g[0:].lower())
            a += (s[0:].lower()+g[0:].lower())
            b += (s[0:1].lower()+g[0:].lower())
            c += (s[0:].lower()+g[0:1].lower())
            d += (g[0:1].lower()+s[0:].lower())
            e += (g[0:].lower()+s[0:1].lower())
            a1 += (s[0:].lower()+'-'+g[0:].lower())
            a2 += (s[0:1].lower()+'-'+g[0:].lower())
            a3 += (s[0:].lower()+'-'+g[0:1].lower())
            a4 += (g[0:1].lower()+'-'+s[0:].lower())
            a5 += (g[0:].lower()+'-'+s[0:1].lower())
          
      
        return (new+'@'+Company, second+'@'+Company, third+'@'+Company, fourth+'@'+Company, fifth+'@'+Company, sixth+'@'+Company, seventh+'@'+Company, a+'@'+Company, b+'@'+Company, c+'@'+Company, d+'@'+Company, e+'@'+Company,a1+'@'+Company,a2+'@'+Company,a3+'@'+Company,a4+'@'+Company,a5+'@'+Company+com)

    def name4(s,g,company):

        l = s.split()
        g = g.split()
        new = ""
        second = "" 
        third = ""
        fourth = ""
        fifth = ""
        sixth = ""
        seventh = ""
        a= ""
        b= ""
        c= ""
        d= ""
        e= ""
        a1= ""
        a2= ""
        a3= ""
        a4= ""
        a5= ""
        a11= ""
        a22= ""
        a33= ""
        a44= ""
        a55= ""
  
        for i in range(len(l)):
            s = l[i]
            g = g[i]
            new += (s[0:].lower())
            second += (s[0:].lower()+'.'+g[0:].lower())
            third += (s[0:1].lower()+'.'+g[0:].lower())
            fourth += (s[0:].lower()+'.'+g[0:1].lower())
            fifth += (g[0:1].lower()+'.'+s[0:].lower())
            sixth += (g[0:].lower()+'.'+s[0:1].lower())
            seventh += (g[0:].lower())
            a += (s[0:].lower()+g[0:].lower())
            b += (s[0:1].lower()+g[0:].lower())
            c += (s[0:].lower()+g[0:1].lower())
            d += (g[0:1].lower()+s[0:].lower())
            e += (g[0:].lower()+s[0:1].lower())
            a1 += (s[0:].lower()+'-'+g[0:].lower())
            a2 += (s[0:1].lower()+'-'+g[0:].lower())
            a3 += (s[0:].lower()+'-'+g[0:1].lower())
            a4 += (g[0:1].lower()+'-'+s[0:].lower())
            a5 += (g[0:].lower()+'-'+s[0:1].lower())
            a11 += (s[0:].lower()+'_'+g[0:].lower())
            a22 += (s[0:1].lower()+'_'+g[0:].lower())
            a33 += (s[0:].lower()+'_'+g[0:1].lower())
            a44 += (g[0:1].lower()+'_'+s[0:].lower())
            a55 += (g[0:].lower()+'_'+s[0:1].lower())
          
      
        return (new+company, second+company, third+company, fourth+company, fifth+company, sixth+company, seventh+company, a+company, b+company, c+company, d+company, e+company,a1+company,a2+company,a3+company,a4+company,a5+company,a11+company,a22+company,a33+company,a44+company,a55+company)




    

    very_high = name4(First_name,Last_name,Company)

    scale = very_high

    a = email_checker(scale)

    wb.sheets[0].range('E2').value = a

def collect_reddits():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    from psaw import PushshiftAPI
    import pandas as pd
    import datetime as dt

    api = PushshiftAPI()



    import datetime
    posted_after = int(datetime.datetime(2022, 1, 1).timestamp())
    posted_before = int(datetime.datetime(2022, 2, 20).timestamp())

    ###CHANGE LIMIT + SUBREDDIT BELOW###
    query = api.search_submissions(subreddit= wb.sheets[0].range('C26').value, after=wb.sheets[0].range('C27').value, before=wb.sheets[0].range('C28').value, limit=wb.sheets[0].range('C29').value)

    submissions = list()
    for element in query:
        submissions.append(element.d_)
    print(len(submissions))

    import pandas as pd
    df = pd.DataFrame(submissions)

    df = df[['id', 'author', 'created_utc', 'domain','url', 'title', 'score', 'selftext', 'num_comments', 'num_crossposts', 'full_link']]


    df_n = pd.DataFrame()
    df = df.sort_values('num_comments',ascending = False)
    df_n['personal_stories'] = df['title']

    a = df

    wb.sheets[1].range('A1').value = a



def topic_analysis():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    import pandas as pd
    import re
    from gensim import corpora, models, similarities
    import nltk
    from nltk.corpus import stopwords
    import numpy as np
    import pandas as pd
    import sys
    # !{sys.executable} -m spacy download en
    import re, numpy as np, pandas as pd
    from pprint import pprint

    import gensim, spacy, logging, warnings
    nlp = spacy.load('en_core_web_sm')
    import gensim.corpora as corpora
    from gensim.utils import  simple_preprocess
    from gensim.models import CoherenceModel
    import matplotlib.pyplot as plt

    from nltk.corpus import stopwords
    stop_words = stopwords.words('english')
    stop_words.extend(['from', 'subject', 're', 'edu', 'use', 'not', 'would', 'say', 'could', '_', 'be', 'know', 'good', 'go', 'get', 'do', 'done', 'try', 'many', 'some', 'nice', 'thank', 'think', 'see', 'rather', 'easy', 'easily', 'lot', 'lack', 'make', 'want', 'seem', 'run', 'need', 'even', 'right', 'line', 'even', 'also', 'may', 'take', 'come'])

    warnings.filterwarnings("ignore",category=DeprecationWarning)
    logging.basicConfig(format='%(asctime)s : %(levelname)s : %(message)s', level=logging.ERROR)


    red_it = pd.read_excel('/Users/willcarrington/demo/people_finder.xlsm','reddits')
    red_it['personal_stories'] = red_it['selftext']
    red_it = pd.DataFrame(red_it['personal_stories'])

#---------------------------------------------------------------------------------

    import gensim
    from gensim.utils import simple_preprocess
    import nltk
    nltk.download('stopwords')
    from nltk.corpus import stopwords
    stop_words = stopwords.words('english')
    stop_words.extend(['from', 'subject', 're', 'edu', 'use','www','https','com','fags','string','sleep','apnea','cpap','co','machine','get','getting'])
    def sent_to_words(sentences):
        for sentence in sentences:
            yield(gensim.utils.simple_preprocess(str(sentence), deacc=True))
    def remove_stopwords(texts):
        return [[word for word in simple_preprocess(str(doc)) 
                 if word not in stop_words] for doc in texts]
    data = red_it.personal_stories.values.tolist()
    data_words = list(sent_to_words(data))
    data_words = remove_stopwords(data_words)

    import gensim.corpora as corpora
    id2word = corpora.Dictionary(data_words)
    texts = data_words
    corpus = [id2word.doc2bow(text) for text in texts]

    from pprint import pprint
    lda_model = gensim.models.LdaMulticore(corpus=corpus,
                                           id2word=id2word,
                                           num_topics=wb.sheets[0].range('C34').value)
    pprint(lda_model.print_topics())
    doc_lda = lda_model[corpus]


#---------------------------------------------------------------------------------


    def sent_to_words(sentences):
        for sent in sentences:
            sent = gensim.utils.simple_preprocess(str(sent), deacc=True) 
            yield(sent)

    data = red_it.personal_stories.values.tolist()
    data_words = list(sent_to_words(data))
    print(data_words[:1])


    bigram = gensim.models.Phrases(data_words, min_count=5, threshold=100) # higher threshold fewer phrases.
    trigram = gensim.models.Phrases(bigram[data_words], threshold=100)  
    bigram_mod = gensim.models.phrases.Phraser(bigram)
    trigram_mod = gensim.models.phrases.Phraser(trigram)


    def process_words(texts, stop_words=stop_words, allowed_postags=['NOUN', 'ADJ', 'VERB', 'ADV']):
        """Remove Stopwords, Form Bigrams, Trigrams and Lemmatization"""
        texts = [[word for word in simple_preprocess(str(doc)) if word not in stop_words] for doc in texts]
        texts = [bigram_mod[doc] for doc in texts]
        texts = [trigram_mod[bigram_mod[doc]] for doc in texts]
        texts_out = []
        for sent in texts:
            doc = nlp(" ".join(sent)) 
            texts_out.append([token.lemma_ for token in doc if token.pos_ in allowed_postags])
        texts_out = [[word for word in simple_preprocess(str(doc)) if word not in stop_words] for doc in texts_out]    
        return texts_out

    data_ready = process_words(data_words)  


#---------------------------------------------------------------------------------

    id2word = corpora.Dictionary(data_ready)

    corpus = [id2word.doc2bow(text) for text in data_ready]

    lda_model = gensim.models.ldamodel.LdaModel(corpus=corpus,
                                               id2word=id2word,
                                               num_topics=wb.sheets[0].range('C34').value, 
                                               random_state=100,
                                               update_every=1,
                                               chunksize=10,
                                               passes=10,
                                               alpha='symmetric',
                                               iterations=100,
                                               per_word_topics=True)


#---------------------------------------------------------------------------------



    def format_topics_sentences(ldamodel=None, corpus=corpus, texts=data):
        sent_topics_df = pd.DataFrame()

        for i, row_list in enumerate(ldamodel[corpus]):
            row = row_list[0] if ldamodel.per_word_topics else row_list            
            row = sorted(row, key=lambda x: (x[1]), reverse=True)
            for j, (topic_num, prop_topic) in enumerate(row):
                if j == 0:
                    wp = ldamodel.show_topic(topic_num)
                    topic_keywords = ", ".join([word for word, prop in wp])
                    sent_topics_df = sent_topics_df.append(pd.Series([int(topic_num), round(prop_topic,4), topic_keywords]), ignore_index=True)
                else:
                    break
        sent_topics_df.columns = ['Dominant_Topic', 'Perc_Contribution', 'Topic_Keywords']

        contents = pd.Series(texts)
        sent_topics_df = pd.concat([sent_topics_df, contents], axis=1)
        return(sent_topics_df)


    df_topic_sents_keywords = format_topics_sentences(ldamodel=lda_model, corpus=corpus, texts=data_ready)

    df_dominant_topic = df_topic_sents_keywords.reset_index()
    df_dominant_topic.columns = ['Document_No', 'Dominant_Topic', 'Topic_Perc_Contrib', 'Keywords', 'Text']
    df_dominant_topic.head(10)



    pd.options.display.max_colwidth = 100

    sent_topics_sorteddf_mallet = pd.DataFrame()
    sent_topics_outdf_grpd = df_topic_sents_keywords.groupby('Dominant_Topic')

    for i, grp in sent_topics_outdf_grpd:
        sent_topics_sorteddf_mallet = pd.concat([sent_topics_sorteddf_mallet, 
                                                 grp.sort_values(['Perc_Contribution'], ascending=False).head(1)], 
                                                axis=0)

    sent_topics_sorteddf_mallet.reset_index(drop=True, inplace=True)

    sent_topics_sorteddf_mallet.columns = ['Topic_Num', "Topic_Perc_Contrib", "Keywords", "Representative Text"]

    sent_topics_sorteddf_mallet.head(10)

    books_read = sent_topics_sorteddf_mallet

#---------------------------------------------------------------------------------

    from matplotlib import pyplot as plt
    from wordcloud import WordCloud, STOPWORDS
    import matplotlib.colors as mcolors

    cols = [color for name, color in mcolors.TABLEAU_COLORS.items()]  # more colors: 'mcolors.XKCD_COLORS'

    cloud = WordCloud(stopwords=stop_words,
                      background_color='white',
                      width=2500,
                      height=1800,
                      max_words=20,
                      colormap='tab10',
                      color_func=lambda *args, **kwargs: cols[i],
                      prefer_horizontal=1.0)

    topics = lda_model.show_topics(formatted=False)

    fig, axes = plt.subplots(4, 4, figsize=(20,20), sharex=True, sharey=True)

    for i, ax in enumerate(axes.flatten()):
        fig.add_subplot(ax)
        topic_words = dict(topics[i][1])
        cloud.generate_from_frequencies(topic_words, max_font_size=300)
        plt.gca().imshow(cloud)
        plt.gca().set_title('Topic ' + str(i), fontdict=dict(size=16))
        plt.gca().axis('off')


    plt.subplots_adjust(wspace=0, hspace=0)
    plt.axis('off')
    plt.margins(x=0, y=0)
    plt.tight_layout()
    wb.sheets[3].pictures.add(fig, name='MyPlot', update=True)
    plt.show()

#---------------------------------------------------------------------------------

    from collections import Counter
    topics = lda_model.show_topics(formatted=False)
    data_flat = [w for w_list in data_ready for w in w_list]
    counter = Counter(data_flat)

    out = []
    for i, topic in topics:
        for word, weight in topic:
            out.append([word, i , weight, counter[word]])

    df = pd.DataFrame(out, columns=['word', 'topic_id', 'importance', 'word_count'])        

    fig, axes = plt.subplots(4, 4, figsize=(20,20), sharey=True, dpi=160)
    cols = [color for name, color in mcolors.TABLEAU_COLORS.items()]
    for i, ax in enumerate(axes.flatten()):
        ax.bar(x='word', height="word_count", data=df.loc[df.topic_id==i, :], color=cols[i], width=0.5, alpha=0.3, label='Word Count')
        ax_twin = ax.twinx()
        ax_twin.bar(x='word', height="importance", data=df.loc[df.topic_id==i, :], color=cols[i], width=0.2, label='Weights')
        ax.set_ylabel('Word Count', color=cols[i])
        ax_twin.set_ylim(0, 0.030); ax.set_ylim(0, 3500)
        ax.set_title('Topic: ' + str(i), color=cols[i], fontsize=16)
        ax.tick_params(axis='y', left=False)
        ax.set_xticklabels(df.loc[df.topic_id==i, 'word'], rotation=30, horizontalalignment= 'right')
        ax.legend(loc='upper left'); ax_twin.legend(loc='upper right')

    fig.tight_layout(w_pad=2)    
    fig.suptitle('Word Count and Importance of Topic Keywords', fontsize=22, y=1.05)   
    wb.sheets[3].pictures.add(fig, name='MyPlot1', update=True)
    plt.show()


def get_tweets():



    wb = xw.Book.caller()
    sheet = wb.sheets[0]



    import os
    import tweepy as tw
    import pandas as pd

    consumer_key= 'IYFL8HKYtOtyusvx0jfygldXP'
    consumer_secret= 'v9NWt16yJaF0PbDR6a0Tbag1RBh8Ul46CTW0fMkuU6OSOS3xEd'
    access_token= '1343573569796767745-GU5uV96mWsSL9SPv3AMekT4SaLtdxL'
    access_token_secret= 'rIZLaRGHqrRzLTfW4k6alieHI0ajaDbZ4PG2vvHtclPeu'


    auth = tw.OAuthHandler(consumer_key, consumer_secret)
    auth.set_access_token(access_token, access_token_secret)
    api = tw.API(auth, wait_on_rate_limit=True)

    search_words = wb.sheets[0].range('F26').value
    search = search_words + " -filter:retweets"


    tweets = tw.Cursor(api.search,
                       q=search,
                       lang="en",
                       since=wb.sheets[0].range('F27').value).items(wb.sheets[0].range('F29').value)

    all_tweets = [[tweet.text, tweet.id, str(tweet.created_at.date()),tweet.user.location,tweet.entities] for tweet in tweets]

    df = pd.DataFrame(all_tweets)

    wb.sheets[2].range('A1').value = df

#------------------------------------------------------------------------------


def topic_analysis_twit():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    import pandas as pd
    import re
    from gensim import corpora, models, similarities
    import nltk
    from nltk.corpus import stopwords
    import numpy as np
    import pandas as pd
    import sys
    # !{sys.executable} -m spacy download en
    import re, numpy as np, pandas as pd
    from pprint import pprint

    import gensim, spacy, logging, warnings
    nlp = spacy.load('en_core_web_sm')
    import gensim.corpora as corpora
    from gensim.utils import  simple_preprocess
    from gensim.models import CoherenceModel
    import matplotlib.pyplot as plt

    from nltk.corpus import stopwords
    stop_words = stopwords.words('english')
    stop_words.extend(['from', 'subject', 're', 'edu', 'use', 'not', 'would', 'say', 'could', '_', 'be', 'know', 'good', 'go', 'get', 'do', 'done', 'try', 'many', 'some', 'nice', 'thank', 'think', 'see', 'rather', 'easy', 'easily', 'lot', 'lack', 'make', 'want', 'seem', 'run', 'need', 'even', 'right', 'line', 'even', 'also', 'may', 'take', 'come'])

    warnings.filterwarnings("ignore",category=DeprecationWarning)
    logging.basicConfig(format='%(asctime)s : %(levelname)s : %(message)s', level=logging.ERROR)


    red_it = pd.read_excel('/Users/willcarrington/demo/people_finder.xlsm','tweets')
    red_it['personal_stories'] = red_it['selftext']
    red_it = pd.DataFrame(red_it['personal_stories'])

#---------------------------------------------------------------------------------

    import gensim
    from gensim.utils import simple_preprocess
    import nltk
    nltk.download('stopwords')
    from nltk.corpus import stopwords
    stop_words = stopwords.words('english')
    stop_words.extend(['from', 'subject', 're', 'edu', 'use','www','https','com','fags','string','sleep','apnea','cpap','co','machine','get','getting'])
    def sent_to_words(sentences):
        for sentence in sentences:
            yield(gensim.utils.simple_preprocess(str(sentence), deacc=True))
    def remove_stopwords(texts):
        return [[word for word in simple_preprocess(str(doc)) 
                 if word not in stop_words] for doc in texts]
    data = red_it.personal_stories.values.tolist()
    data_words = list(sent_to_words(data))
    data_words = remove_stopwords(data_words)

    import gensim.corpora as corpora
    id2word = corpora.Dictionary(data_words)
    texts = data_words
    corpus = [id2word.doc2bow(text) for text in texts]

    from pprint import pprint
    lda_model = gensim.models.LdaMulticore(corpus=corpus,
                                           id2word=id2word,
                                           num_topics=wb.sheets[0].range('F34').value)
    pprint(lda_model.print_topics())
    doc_lda = lda_model[corpus]


#---------------------------------------------------------------------------------


    def sent_to_words(sentences):
        for sent in sentences:
            sent = gensim.utils.simple_preprocess(str(sent), deacc=True) 
            yield(sent)

    data = red_it.personal_stories.values.tolist()
    data_words = list(sent_to_words(data))
    print(data_words[:1])


    bigram = gensim.models.Phrases(data_words, min_count=5, threshold=100) # higher threshold fewer phrases.
    trigram = gensim.models.Phrases(bigram[data_words], threshold=100)  
    bigram_mod = gensim.models.phrases.Phraser(bigram)
    trigram_mod = gensim.models.phrases.Phraser(trigram)


    def process_words(texts, stop_words=stop_words, allowed_postags=['NOUN', 'ADJ', 'VERB', 'ADV']):
        """Remove Stopwords, Form Bigrams, Trigrams and Lemmatization"""
        texts = [[word for word in simple_preprocess(str(doc)) if word not in stop_words] for doc in texts]
        texts = [bigram_mod[doc] for doc in texts]
        texts = [trigram_mod[bigram_mod[doc]] for doc in texts]
        texts_out = []
        for sent in texts:
            doc = nlp(" ".join(sent)) 
            texts_out.append([token.lemma_ for token in doc if token.pos_ in allowed_postags])
        texts_out = [[word for word in simple_preprocess(str(doc)) if word not in stop_words] for doc in texts_out]    
        return texts_out

    data_ready = process_words(data_words)  


#---------------------------------------------------------------------------------

    id2word = corpora.Dictionary(data_ready)

    corpus = [id2word.doc2bow(text) for text in data_ready]

    lda_model = gensim.models.ldamodel.LdaModel(corpus=corpus,
                                               id2word=id2word,
                                               num_topics=wb.sheets[0].range('F34').value, 
                                               random_state=100,
                                               update_every=1,
                                               chunksize=10,
                                               passes=10,
                                               alpha='symmetric',
                                               iterations=100,
                                               per_word_topics=True)


#---------------------------------------------------------------------------------



    def format_topics_sentences(ldamodel=None, corpus=corpus, texts=data):
        sent_topics_df = pd.DataFrame()

        for i, row_list in enumerate(ldamodel[corpus]):
            row = row_list[0] if ldamodel.per_word_topics else row_list            
            row = sorted(row, key=lambda x: (x[1]), reverse=True)
            for j, (topic_num, prop_topic) in enumerate(row):
                if j == 0:
                    wp = ldamodel.show_topic(topic_num)
                    topic_keywords = ", ".join([word for word, prop in wp])
                    sent_topics_df = sent_topics_df.append(pd.Series([int(topic_num), round(prop_topic,4), topic_keywords]), ignore_index=True)
                else:
                    break
        sent_topics_df.columns = ['Dominant_Topic', 'Perc_Contribution', 'Topic_Keywords']

        contents = pd.Series(texts)
        sent_topics_df = pd.concat([sent_topics_df, contents], axis=1)
        return(sent_topics_df)


    df_topic_sents_keywords = format_topics_sentences(ldamodel=lda_model, corpus=corpus, texts=data_ready)

    df_dominant_topic = df_topic_sents_keywords.reset_index()
    df_dominant_topic.columns = ['Document_No', 'Dominant_Topic', 'Topic_Perc_Contrib', 'Keywords', 'Text']
    df_dominant_topic.head(10)



    pd.options.display.max_colwidth = 100

    sent_topics_sorteddf_mallet = pd.DataFrame()
    sent_topics_outdf_grpd = df_topic_sents_keywords.groupby('Dominant_Topic')

    for i, grp in sent_topics_outdf_grpd:
        sent_topics_sorteddf_mallet = pd.concat([sent_topics_sorteddf_mallet, 
                                                 grp.sort_values(['Perc_Contribution'], ascending=False).head(1)], 
                                                axis=0)

    sent_topics_sorteddf_mallet.reset_index(drop=True, inplace=True)

    sent_topics_sorteddf_mallet.columns = ['Topic_Num', "Topic_Perc_Contrib", "Keywords", "Representative Text"]

    sent_topics_sorteddf_mallet.head(10)

    books_read = sent_topics_sorteddf_mallet

#---------------------------------------------------------------------------------

    from matplotlib import pyplot as plt
    from wordcloud import WordCloud, STOPWORDS
    import matplotlib.colors as mcolors

    cols = [color for name, color in mcolors.TABLEAU_COLORS.items()]  # more colors: 'mcolors.XKCD_COLORS'

    cloud = WordCloud(stopwords=stop_words,
                      background_color='white',
                      width=2500,
                      height=1800,
                      max_words=20,
                      colormap='tab10',
                      color_func=lambda *args, **kwargs: cols[i],
                      prefer_horizontal=1.0)

    topics = lda_model.show_topics(formatted=False)

    fig, axes = plt.subplots(4, 2, figsize=(20,20), sharex=True, sharey=True)

    for i, ax in enumerate(axes.flatten()):
        fig.add_subplot(ax)
        topic_words = dict(topics[i][1])
        cloud.generate_from_frequencies(topic_words, max_font_size=300)
        plt.gca().imshow(cloud)
        plt.gca().set_title('Topic ' + str(i), fontdict=dict(size=16))
        plt.gca().axis('off')


    plt.subplots_adjust(wspace=0, hspace=0)
    plt.axis('off')
    plt.margins(x=0, y=0)
    plt.tight_layout()
    wb.sheets[4].pictures.add(fig, name='MyPlot3', update=True)
    plt.show()

#---------------------------------------------------------------------------------

    from collections import Counter
    topics = lda_model.show_topics(formatted=False)
    data_flat = [w for w_list in data_ready for w in w_list]
    counter = Counter(data_flat)

    out = []
    for i, topic in topics:
        for word, weight in topic:
            out.append([word, i , weight, counter[word]])

    df = pd.DataFrame(out, columns=['word', 'topic_id', 'importance', 'word_count'])        

    fig, axes = plt.subplots(4, 2, figsize=(20,20), sharey=True, dpi=160)
    cols = [color for name, color in mcolors.TABLEAU_COLORS.items()]
    for i, ax in enumerate(axes.flatten()):
        ax.bar(x='word', height="word_count", data=df.loc[df.topic_id==i, :], color=cols[i], width=0.5, alpha=0.3, label='Word Count')
        ax_twin = ax.twinx()
        ax_twin.bar(x='word', height="importance", data=df.loc[df.topic_id==i, :], color=cols[i], width=0.2, label='Weights')
        ax.set_ylabel('Word Count', color=cols[i])
        ax_twin.set_ylim(0, 0.030); ax.set_ylim(0, 3500)
        ax.set_title('Topic: ' + str(i), color=cols[i], fontsize=16)
        ax.tick_params(axis='y', left=False)
        ax.set_xticklabels(df.loc[df.topic_id==i, 'word'], rotation=30, horizontalalignment= 'right')
        ax.legend(loc='upper left'); ax_twin.legend(loc='upper right')

    fig.tight_layout(w_pad=2)    
    fig.suptitle('Word Count and Importance of Topic Keywords', fontsize=22, y=1.05)   
    wb.sheets[4].pictures.add(fig, name='MyPlot4', update=True)
    plt.show()

