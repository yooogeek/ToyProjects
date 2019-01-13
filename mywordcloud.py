from wordcloud import WordCloud
import matplotlib.pyplot as plt

filename = "test1.txt"
with open(filename, encoding='UTF-8') as f:
    mytext = f.read()
mywordcloud = WordCloud(font_path='C:/Windows/Fonts/STKAITI.TTF').generate(mytext)
plt.imshow(mywordcloud, interpolation='bilinear')
plt.axis("off")
plt.show()