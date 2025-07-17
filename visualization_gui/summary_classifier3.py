import pandas as pd
import math
from gensim.corpora.dictionary import Dictionary
from gensim.models import LdaModel, TfidfModel
from text_preprocessing import process_text


# Load Japanese stopwords
def load_jp_stopwords(path="src/slothlib.txt"):
    with open(path, encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]

def main(input_excel, output_excel, summary_col='Summary', n_topics=8):
    # Load data
    df = pd.read_excel(input_excel)
    # Drop rows where summary_col is NaN
    df = df.dropna(subset=[summary_col])
    df[summary_col] = df[summary_col].astype(str)
    # Preprocess summaries
    processed = process_text(df[summary_col])
    df['tokens'] = processed
    # Prepare texts for gensim
    texts = [t.split() for t in df['tokens'].values]
    dictionary = Dictionary(texts)
    dictionary.filter_extremes(no_below=3, no_above=0.8)
    corpus = [dictionary.doc2bow(text) for text in texts]
    # TFIDF model
    def new_idf(docfreq, totaldocs, log_base=2.0, add=1.0):
        return add + math.log(1.0 * totaldocs / docfreq, log_base)
    tfidf_model = TfidfModel(corpus, wglobal=new_idf, normalize=False)
    tfidf_corpus = tfidf_model[corpus]
    # LDA model
    lda = LdaModel(corpus=tfidf_corpus, num_topics=n_topics, id2word=dictionary, random_state=0)
    # Assign topics
    topic_assignments = []
    for bow in tfidf_corpus:
        topic_probs = lda.get_document_topics(bow)
        topic = max(topic_probs, key=lambda x: x[1])[0]
        topic_assignments.append(topic)
    # Get topic names (top words)
    topic_names = []
    for i in range(n_topics):
        words = [dictionary[w] for w, _ in lda.get_topic_terms(i, topn=3)]
        topic_names.append('Topic{}_{}'.format(i+1, '_'.join(words)))
    # Group rows by topic
    topic_to_rows = {name: [] for name in topic_names}
    for idx, topic_idx in zip(df.index, topic_assignments):
        topic_name = topic_names[topic_idx]
        topic_to_rows[topic_name].append(df.loc[idx])
    # Write to Excel with multiple sheets, excluding 'tokens' column
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        for topic, rows in topic_to_rows.items():
            df_out = pd.DataFrame(rows)
            if 'tokens' in df_out.columns:
                df_out = df_out.drop(columns=['tokens'])
            df_out.to_excel(writer, sheet_name=topic[:31], index=False)

if __name__ == "__main__":
    main(
        input_excel="mask.xlsx",  # Change to your input file
        output_excel="categorized_defects.xlsx",  # Change to your desired output file
        summary_col="Summary",
        n_topics=8   # You can adjust the number of categories
    )
