from HNLP import Analyzer;
import sys

if __name__ == "__main__":

    file_name = sys.argv[1]
    tagger = sys.argv[2]
    analyzer = Analyzer(name= file_name,tagger=tagger)
    analyzer.analyze()





