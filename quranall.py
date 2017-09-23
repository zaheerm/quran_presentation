import quranpptx


if __name__ == '__main__':
    suras = quranpptx.load_suras("quran-uthmani.xml", "shakir_table.csv")
    for num, name in quranpptx.SURA_NAMES.items():
        quranpptx.create_presentation(suras, num, "{}_{}.pptx".format(num, name), arabic_font='Arial')
