import pandas as pd
import spacy
from collections import Counter
import en_core_web_sm
from yake import KeywordExtractor
from spacy.tokens import Doc

Doc.set_extension("yake", default=None, force=True)  # force=True

# 加载英文模型
nlp = spacy.load("en_core_web_sm")
print("Load ok")

# 自定义关键短语提取组件
@spacy.language.Language.component("yake_extractor")
def extract_key_phrases(doc):
    extractor = KeywordExtractor()
    key_phrases = extractor.extract_keywords(doc.text)
    doc._.yake = key_phrases
    return doc


# 添加自定义关键短语提取组件
nlp.add_pipe("yake_extractor", last=True)

# 文本预处理函数：转换为小写，移除停用词和标点
def preprocess_text(text):
    doc = nlp(text.lower())
    result = [token.text for token in doc if not token.is_stop and not token.is_punct]
    return result

# 读取Excel文件
file_path = 'reviews_all.xlsx'  # 替换为你的文件路径
df = pd.read_excel(file_path)

# 假设评论在名为"comments"的列中，根据你的文件实际情况调整列名
comments = df['comments']

# 应用文本预处理
processed_texts = [preprocess_text(text) for text in comments]

# 扁平化处理后文本列表并计算词频
all_words = [word for sublist in processed_texts for word in sublist]
word_freq = Counter(all_words)

# 获取频率最高的前20个词
top_20_words = word_freq.most_common(20)

# 打印结果
print("Top 20 Keywords and their Frequencies:")
for word, freq in top_20_words:
    print(f"{word}: {freq}")


# 提取关键短语
key_phrases = []
for doc in nlp.pipe(comments):
    key_phrases.extend(doc._.yake)

# 计算关键短语频率
phrase_freq = Counter(key_phrases)
print("\nTop 20 Key Phrases and their Frequencies:")
for phrase, freq in phrase_freq.most_common(20):
    print(f"{phrase}: {freq}")