import MeCab
import ipadic
import re
import mojimoji
from sklearn.feature_extraction.text import ENGLISH_STOP_WORDS
from nltk.stem import WordNetLemmatizer

# Load Japanese stopwords
def load_jp_stopwords(path="src/slothlib.txt"):
    with open(path, encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]

# MeCab tokenizer: keep nouns, adjectives, proper nouns
def process_text(series):
    stop_words = set(load_jp_stopwords())
    english_stopwords = set(ENGLISH_STOP_WORDS)
    mecab = MeCab.Tagger(ipadic.MECAB_ARGS)
    lemmatizer = WordNetLemmatizer()
    def tokenizer_func(text):
        tokens = []
        # Japanese token extraction
        node = mecab.parseToNode(str(text))
        while node:
            features = node.feature.split(',')
            length = len(features)
            if length > 6:
                surface = features[6]
            else:
                surface = node.surface
            is_pass = (surface == '*') or (len(surface) < 2) or (surface in stop_words)
            noun_flag = (features[0] == '名詞')
            proper_noun_flag = (features[0] == '名詞') and (features[1] == '固有名詞')
            adjective_flag = (features[0] == '形容詞') and (features[1] == '自立')
            if is_pass:
                node = node.next
                continue
            elif proper_noun_flag:
                tokens.append(surface)
            elif noun_flag:
                tokens.append(surface)
            elif adjective_flag:
                tokens.append(surface)
            node = node.next
        # English token extraction with lemmatization (nouns and verbs)
        english_tokens = re.findall(r'[a-zA-Z]{2,}', text)
        lemmatized_english = []
        for w in english_tokens:
            w_lower = w.lower()
            lemma_noun = lemmatizer.lemmatize(w_lower, pos='n')
            lemma_verb = lemmatizer.lemmatize(w_lower, pos='v')
            # Prefer verb lemma if it changes the word, else noun lemma
            if lemma_verb != w_lower:
                lemma = lemma_verb
            else:
                lemma = lemma_noun
            lemmatized_english.append(lemma)
        tokens.extend([w for w in lemmatized_english if w not in english_stopwords])
        print('Extracted tokens: ', tokens)
        return " ".join(tokens)
    series = series.map(tokenizer_func)
    series = series.map(lambda x: x.lower())
    series = series.map(mojimoji.zen_to_han)
    return series
