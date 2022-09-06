import pandas as pd
import seaborn as sb
import matplotlib as mp
import yfinance as yf

# Extract the symbols of Dow 30 from Excel file
obj = pd.read_excel('input.xlsx').values.tolist()
target = []
for i in range(0, len(obj)):
    target.append(obj[i][0])

# Download close prices of the stocks
r = yf.download(target, start='2022-01-02', end='2022-07-01')['Close']

# Convert data frame into a line graph
p = sb.lineplot(data = r.apply(lambda row: row / r.iloc[[0]].values[0] - 1, axis = 1), dashes = False, palette = 'tab10')
p.set(xlabel = None, title = 'Trends of DOW Jones Industrial Components (DOW 30) \nbetween January and June in 2022')
mp.pyplot.xticks(rotation = 30)
p.yaxis.set_major_formatter(mp.ticker.FuncFormatter(lambda y, _: '{:.1%}'.format(y)))

# Return a line graph and an Excel file with all the close prices among Dow 30
p.legend(fontsize = 6, bbox_to_anchor = (1, 1)).get_figure().savefig("output.png")
r.to_excel('output.xlsx')
