#%%
#imports
from fuzzywuzzy import fuzz,process
import fuzzymatcher
import pandas as pd
# %%
fuzz.WRatio('reeding','Reading')
# %%
# data import
#df1 = pd.read_excel('Textbook projectSummerFall2021-2.xlsx',header=2)
df1 = pd.read_excel('Textbook projectSummerFall2021-2.xlsx',header=2,sheet_name='SpringAlmaOutput')
#df1 = pd.read_excel('Textbook projectSummerFall2021-2.xlsx',header=2,sheet_name='FallAlmaOutput')
df1.head()
# %%
#df2 = pd.read_excel('Textbook projectSummerFall2021-2.xlsx',sheet_name='CSApprovedAdoptionList')
df2 = pd.read_excel('Textbook projectSummerFall2021-2.xlsx',sheet_name='SpringBookstoreList')
#df2 = pd.read_excel('Textbook projectSummerFall2021-2.xlsx',sheet_name='FallBookstoreList')
df2.head()
#%% 
#method 1
matched_results = fuzzymatcher.fuzzy_left_join(df1,
                                            df2,
                                            'Title',
                                            'Long Title')

#%%
matched_results[['best_match_score','Title','Long Title','Internal ID']].to_excel('fuzzymatcherresults_min_spring.xlsx')
matched_results.to_excel('fuzzymatcherresults_full_spring.xlsx')
# %%
#method 2
import recordlinkage as rl
from recordlinkage.index import Full
# %%
indexer = rl.Index()
indexer.add(Full())

pairs = indexer.index(df1,df2,)
print(len(pairs))

comparer = rl.Compare()
comparer.string('Title','Long Title',threshold=0.55,label='Title')

potential_matches = comparer.compute(pairs, df1,df2)
matches = potential_matches[potential_matches.sum(axis=1)> 0].reset_index()
#print(matches)

accumulated = matches.loc[:,['level_0','level_1']].merge(df1.loc[:,['Title','ISBN']], left_on='level_0',right_index=True)
accumulated = accumulated.merge(df2.loc[:,['Long Title','Internal ID']], left_on='level_1',right_index=True)
accumulated.head()
accumulated.to_excel('recordlinkagemethod_55_spring.xlsx')
# %%
