import glob
import pathlib

from beautify_table.beautify import *


def to_float(s):
    try:
        return float(s)
    except ValueError:
        return s


def dataframe_to_couples(df):
    dfd = df.to_dict()
    lst = []
    for k, v in dfd.items():
        kl = k.lower().split(" ")
        for ki, vi in v.items():
            kil = ki.lower().split(" ")
            lst.append((set(kl).union(set(kil)), to_float(vi), k + " " + ki))
    return lst


def mapping_table_to_words(df):
    transform = lambda x: [x[c].lower().split(" ") if isinstance(x[c], str) else "" for c in x.index]
    df["words"] = df.apply(transform, axis=1)
    df["words"] = df.apply(lambda x: set(b for a in x["words"] for b in a), axis=1)
    return df


def mapping_match(words, mapped):
    return len(mapped.intersection(set(words)))


df = pd.read_excel(f"../data/mapping_table.xlsx", header=0)
words = mapping_table_to_words(df)
words["best_match_val"] = -1
words["match"] = ""
words["mapping"] = ""
# display(words)

for i, f in enumerate(glob.glob("../data/nice/*.xlsx", recursive=True)):
    f = pathlib.Path(f)
    df = pd.read_excel(f, header=0)
    df = df.set_index(df.columns[0])
    #     display(df)
    couples = dataframe_to_couples(df)
    for couple in couples:
        words["mapping_match"] = words.apply(lambda x: mapping_match(x["words"], couple[0]), axis=1)
        best_idx = words["mapping_match"].idxmax()
        best_val = words["mapping_match"].max()
        if best_val > words.loc[best_idx, "best_match_val"] and best_val >= 2:
            words.at[best_idx, "match"] = couple[2]
            words.at[best_idx, "mapping"] = couple[1]
            words.at[best_idx, "best_match_val"] = best_val

del words["words"]
del words["mapping_match"]
words.to_excel("../data/mapping_result.xlsx")


