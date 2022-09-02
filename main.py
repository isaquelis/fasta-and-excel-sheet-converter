def main():
    import converters as cvt

    entry = input("enter 0 to convert .xlsx  to .fasta or\n" +
                  "enter 1 to convert .fasta to  .xlsx\n" +
                  "your entry: ")

    if entry == '0':
        file = input("\nenter xlsx-file directory path: ")

        if file == "":
            # If the input be empty-string it will be select a test-file.
            file = "testFiles/human_sequence.xlsx"
        cvt.fastaOut(file)

    elif entry == '1':
        file = input("\nenter fasta-file directory path: ")

        if file == "":
            # If the input be empty-string it will be select a test-file.
            file = "testFiles/human_sequence.fasta"
        cvt.xlsxOut(file)


if __name__ == "__main__":
    main()
