from konlpy.tag import *
import xlrd
import helpers



class Analyzer:
    def __init__(self,name,tagger):
        print("Analyzing....")
        self.name = name
        self.tagger = tagger


    def analyze(self):
        # morpheme analyzer #
        if (self.tagger == 'Twitter'):
            twit = Twitter()
        else:
            kkma = Kkma()

        # ----------- variables start ----------- #
        self.lineCnt = 0
        self.lineCnt2 = 0
        self.lineNum = 0
        self.sp_cnt1 = 0
        self.sp_cnt2 = 0

        eojeolCnt = 0
        self.total_turn_count = 0
        self.total_speak_count = 0
        self.total_eojeol_count = 0
        # 문장 구성요소

        eojeolTotal = 0
        self.morphTotal = 0

        self.nounCnt2 = 0
        self.pronCnt2 = 0
        verbCnt2 = 0
        self.susaCnt2 = 0
        self.verbCnt2 = 0
        self.adjCnt2 = 0
        self.deterCnt2 = 0
        self.advCnt2 = 0
        self.itjCnt2 = 0
        self.josaCnt2 = 0
        self.eomiCnt2 = 0

        self.nounList2 = {}
        self.pronounList2 = {}
        self.susaList2 = {}
        self.verbList2 = {}
        self.adjList2 = {}
        self.deterList2 = {}
        self.advList2 = {}
        self.interjectList2 = {}
        self.josaList2 = {}
        self.eomiList2 = {}

        self.totalSylCnt = 0
        self.dialogue_list = []
        self.total_natmal_list = {}

        # ----------- variables end ----------- #

        self.file_name = self.name
        self.workbook = xlrd.open_workbook(str(self.file_name) + '.xlsx')

        self.worksheet = self.workbook.sheet_by_index(0)

        self.turn_val = self.worksheet.col_values(0)
        self.speak_val = self.worksheet.col_values(1)
        self.dialogue = self.worksheet.col_values(2)


        for turn in self.turn_val:
            if helpers.isNumber(turn) is True:
                self.total_turn_count += 1

        for speak in self.speak_val:
            if helpers.isNumber(speak) is True:
                self.total_speak_count += 1

        for sentence in self.dialogue:
            if sentence[:1] == '아':
                self.lineCnt += 1
                sentence = helpers.text_cleaner(sentence)
                while True:
                    start = sentence.find('(')
                    end = sentence.find(')')
                    if start >= end: break
                    sentence = sentence[:start] + sentence[end + 1:]
                sentence = ' '.join(sentence.split())  # 중첩 공백 제거
                self.dialogue_list.append(sentence[2:])
            # 문법 요소 Counting


        for line in self.dialogue_list:
            nounCnt = 0
            pronCnt = 0
            susaCnt = 0
            verbCnt = 0
            adjCnt = 0
            deterCnt = 0
            advCnt = 0
            itjCnt = 0
            josaCnt = 0
            eomiCnt = 0
            eojeolCnt = 0
            sylCnt = 0

            if (self.tagger == 'Twitter'):
                pos = twit.pos(line)
            else:
                pos = kkma.pos(line)

            self.lineNum += 1

            for ch in line:
                if ch == '\n' or ch == ' ':
                    continue
                else:
                    sylCnt += 1

            self.sp_cnt2 += 1

            for eojeol in line:
                if eojeol == " ":
                    eojeolCnt += 1

            for token, tag in pos:
                if self.tagger == 'Twitter':
                    noun_flag = False
                    josa_flag= False
                    adj_flag = False
                    verb_flag = False
                    adv_flag = False

                    if tag == 'Noun':
                        for noun in self.nounList2:
                            if noun == token:
                                self.nounList2[token] += 1
                                noun_flag = True
                                break
                        if noun_flag == False:
                            self.nounList2[token] = 1
                        nounCnt +=1
                    if nounCnt >= 2:
                        self.nounCnt2 += 1

                    if tag == 'Josa':
                        for josa in self.josaList2:
                            if josa == token:
                                self.josaList2[token] += 1
                                josa_flag = True
                                break
                        if josa_flag == False:
                            self.josaList2[token] = 1
                        josaCnt += 1
                    if josaCnt >= 2:
                        self.josaCnt2 += 1

                    if tag == 'Adjective':
                        for adj in self.adjList2:
                            if adj == token:
                                self.adjList2[token] += 1
                                adj_flag = True
                                break
                        if adj_flag == False:
                            self.adjList2[token] = 1
                        adjCnt += 1
                    if adjCnt >= 2:
                        self.adjCnt2 += 1

                    if tag == 'Adverb':
                        for adv in self.advList2:
                            if adv == token:
                                self.advList2[token] += 1
                                adv_flag = True
                                break
                        if adv_flag == False:
                            self.advList2[token] = 1
                        advCnt += 1
                    if advCnt >= 2:
                        self.advCnt2 += 1


                    if tag == 'Verb':
                        for verb in self.verbList2:
                            if verb == token:
                                self.verbList2[token] += 1
                                verb_flag = True
                                break
                        if verb_flag == False:
                            self.verbList2[token] = 1
                        verbCnt += 1
                    if verbCnt >= 2:
                        verbCnt2 += 1


                else:

                    n_flag = False
                    pn_flag = False
                    nr_flag = False
                    vv_flag = False
                    va_flag = False
                    md_flag = False
                    ma_flag = False
                    ic_flag = False
                    jk_flag = False
                    ep_flag = False
                    ###############################################################
                    if tag == 'NNG' or tag == 'NNP' or tag == 'NNB' or tag == 'NNM':
                        for noun in self.nounList2:
                            if noun == token:
                                self.nounList2[token] += 1
                                n_flag = True
                                break
                        if n_flag == False:
                            self.nounList2[token] = 1
                        nounCnt += 1
                    if nounCnt >= 2:
                        self.nounCnt2 += 1
                    ###############################################################


                    ###############################################################
                    elif tag == 'NP':
                        for pron in self.pronounList2:
                            if pron == token:
                                self.pronounList2[token] += 1
                                pn_flag = True
                                break
                        if pn_flag == False:
                            self.pronounList2[token] = 1
                        pronCnt += 1
                        if pronCnt >= 2:
                            self.pronCnt2 += 1
                    ###############################################################


                    ###############################################################
                    elif tag == 'VV':
                        for verb in self.verbList2:
                            if verb == token:
                                self.verbList2[token] += 1
                                vv_flag = True
                                break
                        if vv_flag == False:
                            self.verbList2[token] = 1
                        verbCnt += 1
                        if verbCnt >= 2:
                            verbCnt2 += 1
                    ###############################################################


                    ###############################################################
                    elif tag == 'MAG' or tag == 'MAC':
                        for adv in self.advList2:
                            if adv == token:
                                self.advList2[token] += 1
                                ma_flag = True
                                break
                        if ma_flag == False:
                            self.advList2[token] = 1
                        advCnt += 1
                        if advCnt >= 2:
                            self.advCnt2 += 1
                    ###############################################################


                    ###############################################################
                    elif tag == 'VA':
                        for adj in self.adjList2:
                            if adj == token:
                                self.adjList2[token] += 1
                                va_flag = True
                                break
                        if va_flag == False:
                            self.adjList2[token] = 1
                        adjCnt += 1
                        if adjCnt >= 2:
                            self.adjCnt2 += 1
                    ###############################################################


                    ###############################################################
                    elif tag == 'MDT':
                        for deter in self.deterList2:
                            if deter == token:
                                self.deterList2[token] += 1
                                md_flag = True
                                break
                        if md_flag == False:
                            self.deterList2[token] = 1
                        deterCnt += 1
                        if deterCnt >= 2:
                            self.deterCnt2 += 1
                    ###############################################################


                    ###############################################################
                    elif tag == 'IC':
                        for itj in self.interjectList2:
                            if itj == token:
                                self.interjectList2[token] += 1
                                ic_flag = True
                                break
                        if ic_flag == False:
                            self.interjectList2[token] = 1
                        itjCnt += 1
                        if itjCnt >= 2:
                            self.itjCnt2 += 1
                    ###############################################################


                    ###############################################################
                    elif tag == 'NR':
                        for susa in self.susaList2:
                            if susa == token:
                                self.susaList2[token] += 1
                                nr_flag = True
                                break
                        if nr_flag == False:
                            self.susaList2[token] = 1
                        susaCnt += 1
                        if susaCnt >= 2:
                            self.susaCnt2 += 1
                    ###############################################################


                    ###############################################################
                    elif tag == 'JKS' or tag == 'JKC' or tag == 'JKG' or tag == 'JKO' or tag == 'JKM' or tag == 'JKI' or tag == 'JKQ' or tag == 'JC' or tag == 'JX':
                        for josa in self.josaList2:
                            if josa == token:
                                self.josaList2[token] += 1
                                jk_flag = True
                                break
                        if jk_flag == False:
                            self.josaList2[token] = 1
                        josaCnt += 1
                        if josaCnt >= 2:
                            self.josaCnt2 += 1
                    ###############################################################


                    ###############################################################
                    elif tag == 'EPH' or tag == 'EPT' or tag == 'EPP' or tag == 'EFN' or tag == 'EFQ' or tag == 'EFO' or tag == 'EFA' or tag == 'EFI' or tag == 'EFR' or tag == 'ECE' or tag == 'ECD' or tag == 'ECS' or tag == 'ETN' or tag == 'ETD':
                        for eomi in self.eomiList2:
                            if eomi == token:
                                self.eomiList2[token] += 1
                                ep_flag = True
                                break
                        if ep_flag == False:
                            self.eomiList2[token] = 1
                        eomiCnt += 1
                        if eomiCnt >= 2:
                            self.eomiCnt2 += 1
                            ###############################################################

            self.sp_cnt1 += 1
            self.total_eojeol_count += eojeolCnt
            self.totalSylCnt += sylCnt



            # 형태소 개수 Counting
            if self.tagger == 'Twitter':
                morpheme = twit.morphs(line)
            else:
                morpheme = kkma.morphs(line)
            morphCnt = 0
            for morph in morpheme:
                natmal_flag = False
                for natmal in self.total_natmal_list:
                    if natmal == morph:
                        self.total_natmal_list[natmal] += 1
                        natmal_flag = True
                        break
                if natmal_flag == False:
                    self.total_natmal_list[morph] = 1
                morphCnt += 1
            if morphCnt >= 2:
                self.lineCnt2 += 1
            self.morphTotal += morphCnt

        print()
        print()
        MLUm = self.morphTotal / self.total_speak_count  # Mean Length of Utterance , morpheme  / 평균 발화 길이(형태소)
        MLUw = self.total_eojeol_count / self.total_speak_count  # Mean Length of Utterance , morpheme  / 평균 발화 길이(형태소)
        # Mean Length of Utterance , word / 평균 발화 길이(단어)
        MSL = len(self.total_natmal_list) / self.lineCnt2  # Mean Syntactic Length / 평균 구문 길이

        TNW = 0  # 전체 낱말 수
        NDW = 0  # 서로 다른 낱말 수

        for i in self.total_natmal_list:
            NDW += 1
            TNW += self.total_natmal_list[i]
        TTR = TNW / NDW  # 어휘 다양도



        #######################################
        nounTNW = 0
        nounNDW = 0
        for i in self.nounList2:
            nounTNW += self.nounList2[i]  # 명사 TNW
            nounNDW += 1  # 명사 NDW

        if nounTNW is 0:
            nounTNW = 0
            nounNDW = 0
            nounTTR = 0
        else:
            nounTTR = round((nounTNW / nounNDW), 2)

        #######################################

        #######################################
        pronTNW = 0
        pronNDW = 0
        for i in self.pronounList2:
            pronTNW += self.pronounList2[i]
            pronNDW += 1

        if pronTNW is 0:
            pronTNW = 0
            pronNDW = 0
            pronTTR = 0
        else:
            pronTTR = round((pronTNW / pronNDW), 2)
        #######################################

        #######################################
        susaTNW = 0
        susaNDW = 0
        for i in self.susaList2:
            susaTNW += self.susaList2[i]
            susaNDW += 1
        if susaTNW is 0:
            susaTNW = 0
            susaNDW = 0
            susaTTR = 0
        else:
            susaTTR = round((susaTNW / susaNDW), 2)
        #######################################

        #######################################
        verbTNW = 0
        verbNDW = 0
        for i in self.verbList2:
            verbTNW += self.verbList2[i]
            verbNDW += 1

        if verbTNW is 0:
            verbTNW = 0
            verbNDW = 0
            verbTTR = 0
        else:
            verbTTR = round((verbTNW / verbNDW), 2)
        #######################################

        #######################################
        adjTNW = 0
        adjNDW = 0
        for i in self.adjList2:
            adjTNW += self.adjList2[i]
            adjNDW += 1

        if adjTNW is 0:
            adjTNW = 0
            adjNDW = 0
            adjTTR = 0
        else:
            adjTTR = round((adjTNW / adjNDW), 2)

        #######################################

        #######################################
        deterTNW = 0
        deterNDW = 0
        for i in self.deterList2:
            deterTNW += self.deterList2[i]
            deterNDW += 1

        if deterTNW is 0:
            deterTNW = 0
            deterNDW = 0
            deterTTR = 0
        else:
            deterTTR = round((deterTNW / deterNDW), 2)
        #######################################

        #######################################
        advTNW = 0
        advNDW = 0
        for i in self.advList2:
            advTNW += self.advList2[i]
            advNDW += 1
        if advTNW is 0:
            advTNW = 0
            advNDW = 0
            advTTR = 0
        else:
            advTTR = round((advTNW / advNDW), 2)

        #######################################

        #######################################
        itjTNW = 0
        itjNDW = 0
        for i in self.interjectList2:
            itjTNW += self.interjectList2[i]
            itjNDW += 1
        if itjTNW is 0:
            itjTNW = 0
            itjNDW = 0
            itjTTR = 0
        else:
            itjTTR = round((itjTNW / itjNDW), 2)
        #######################################

        #######################################
        josaTNW = 0
        josaNDW = 0
        for i in self.josaList2:
            josaTNW += self.josaList2[i]
            josaNDW += 1
        if josaTNW is 0:
            josaTNW = 0
            josaNDW = 0
            josaTTR = 0
        else:
            josaTTR = round((josaTNW / josaNDW), 2)

        #######################################

        #######################################
        eomiTNW = 0
        eomiNDW = 0
        for i in self.eomiList2:
            eomiTNW += self.eomiList2[i]
            eomiNDW += 1
        if eomiTNW is 0:
            eomiTNW = 0
            eomiNDW = 0
            eomiTTR = 0
        else:
            eomiTTR = round((eomiTNW / eomiNDW), 2)


            #######################################)
        with open(self.name + ' 분석 결과' + '.txt', 'w', encoding='utf-8') as txt:
            txt.write('전사자료 분석정보' + '\n')
            txt.write("\t발화자 : " + self.name + '\n')
            txt.write("\t전사 날짜 : " + '\n')
            txt.write("\t현재 나이 : " + '\n')
            txt.write('\n')

            txt.write("전사자료 길이" + "\n")
            txt.write('\t총 발화 수: ' + str(self.total_speak_count) + '\n')
            txt.write('\n')

            MLUm = self.morphTotal / self.total_speak_count  # Mean Length of Utterance , morpheme  / 평균 발화 길이(형태소)
            MLUw = self.total_eojeol_count / self.total_speak_count  # Mean Length of Utterance , morpheme  / 평균 발화 길이(형태소)
            # Mean Length of Utterance , word / 평균 발화 길이(단어)
            MSL = len(self.total_natmal_list) /self.lineCnt2  # Mean Syntactic Length / 평균 구문 길이

            TNW = 0  # 전체 낱말 수
            NDW = 0  # 서로 다른 낱말 수

            for i in self.total_natmal_list:
                NDW += 1
                TNW += self.total_natmal_list[i]
            TTR = TNW / NDW  # 어휘 다양도

            txt.write('구문부/형태부' + '\n')
            txt.write('\t평균 발화 길이 (형태소, MLUm) : ' + str(round(MLUm, 2)) + '\n')
            txt.write('\t평균 발화 길이 (단어, MLUw) : ' + str(round(MLUw, 2)) + '\n')
            txt.write('\n')

            txt.write('의미부' + '\n')
            txt.write('\t전체 낱말 수(NTW) : ' + str(round(TNW, 2)) + '\n')
            txt.write('\t서로 다른 낱말 수(NDW) : ' + str(round(NDW, 2)) + '\n')
            txt.write('\t어휘 다양도(TTR) : ' + str(round(TTR, 2)) + '\n')
            txt.write('\n')
            txt.write('품사별 측정치' + '\n')

            #######################################
            nounTNW = 0
            nounNDW = 0
            for i in self.nounList2:
                nounTNW += self.nounList2[i]  # 명사 TNW
                nounNDW += 1  # 명사 NDW

            if nounTNW is 0:
                txt.write('\t명사 TNW : ' + str(0) + '\n')
                txt.write('\t명사 NDW : ' + str(0) + '\n')
                txt.write('\t명사 TTR : ' + str(0) + '\n' + '\n')
            else:
                txt.write('\t명사 TNW : ' + str(round(nounTNW, 2)) + '\n')
                txt.write('\t명사 NDW : ' + str(round(nounNDW, 2)) + '\n')
                txt.write('\t명사 TTR : ' + str(round((nounTNW / nounNDW), 2)) + '\n' + '\n')

            #######################################

            #######################################
            pronTNW = 0
            pronNDW = 0
            for i in self.pronounList2:
                pronTNW += self.pronounList2[i]
                pronNDW += 1

            if pronTNW is 0:
                txt.write('\t대명사 TNW : ' + str(0) + '\n')
                txt.write('\t대명사 NDW : ' + str(0) + '\n')
                txt.write('\t대명사 TTR : ' + str(0) + '\n' + '\n')
            else:
                txt.write('\t대명사 TNW : ' + str(round(pronTNW, 2)) + '\n')
                txt.write('\t대명사 NDW : ' + str(round(pronNDW, 2)) + '\n')
                txt.write('\t대명사 TTR : ' + str(round((pronTNW / pronNDW), 2)) + '\n' + '\n')
            #######################################

            #######################################
            susaTNW = 0
            susaNDW = 0
            for i in self.susaList2:
                susaTNW += self.susaList2[i]
                susaNDW += 1
            if susaTNW is 0:
                txt.write('\t수사 TNW : ' + str(0) + '\n')
                txt.write('\t수사 NDW : ' + str(0) + '\n')
                txt.write('\t수사 TTR : ' + str(0) + '\n' + '\n')
            else:
                txt.write('\t수사 TNW : ' + str(round(susaTNW, 2)) + '\n')
                txt.write('\t수사 NDW : ' + str(round(susaNDW, 2)) + '\n')
                txt.write('\t수사 TTR : ' + str(round((susaTNW / susaNDW), 2)) + '\n' + '\n')
            #######################################

            #######################################
            verbTNW = 0
            verbNDW = 0
            for i in self.verbList2:
                verbTNW += self.verbList2[i]
                verbNDW += 1

            if verbTNW is 0:
                txt.write('\t동사 TNW : ' + str(0) + '\n')
                txt.write('\t동사 NDW : ' + str(0) + '\n')
                txt.write('\t동사 TTR : ' + str(0) + '\n' + '\n')
            else:
                txt.write('\t동사 TNW : ' + str(round(verbTNW, 2)) + '\n')
                txt.write('\t동사 NDW : ' + str(round(verbNDW, 2)) + '\n')
                txt.write('\t동사 TTR : ' + str(round((verbTNW / verbNDW), 2)) + '\n' + '\n')
            #######################################

            #######################################
            adjTNW = 0
            adjNDW = 0
            for i in self.adjList2:
                adjTNW += self.adjList2[i]
                adjNDW += 1

            if adjTNW is 0:
                txt.write('\t형용사 TNW : ' + str(0) + '\n')
                txt.write('\t형용사 NDW : ' + str(0) + '\n')
                txt.write('\t형용사 TTR : ' + str(0) + '\n' + '\n')
            else:
                txt.write('\t형용사 TNW : ' + str(round(adjTNW, 2)) + '\n')
                txt.write('\t형용사 NDW : ' + str(round(adjNDW, 2)) + '\n')
                txt.write('\t형용사 TTR : ' + str(round((adjTNW / adjNDW), 2)) + '\n' + '\n')
            #######################################

            #######################################
            deterTNW = 0
            deterNDW = 0
            for i in self.deterList2:
                deterTNW += self.deterList2[i]
                deterNDW += 1

            if deterTNW is 0:
                txt.write('\t관형사 TNW : ' + str(0) + '\n')
                txt.write('\t관형사 NDW : ' + str(0) + '\n')
                txt.write('\t관형사 TTR : ' + str(0) + '\n' + '\n')
            else:
                txt.write('\t관형사 TNW : ' + str(round(deterTNW, 2)) + '\n')
                txt.write('\t관형사 NDW : ' + str(round(deterNDW, 2)) + '\n')
                txt.write('\t관형사 TTR : ' + str(round((deterTNW / deterNDW), 2)) + '\n' + '\n')
            #######################################

            #######################################
            advTNW = 0
            advNDW = 0
            for i in self.advList2:
                advTNW += self.advList2[i]
                advNDW += 1
            if advTNW is 0:
                txt.write('\t부사 TNW : ' + str(0) + '\n')
                txt.write('\t부사 NDW : ' + str(0) + '\n')
                txt.write('\t부사 TTR : ' + str(0) + '\n' + '\n')
            else:
                txt.write('\t부사 TNW : ' + str(round(advTNW, 2)) + '\n')
                txt.write('\t부사 NDW : ' + str(round(advNDW, 2)) + '\n')
                txt.write('\t부사 TTR : ' + str(round((advTNW / advNDW), 2)) + '\n' + '\n')
            #######################################

            #######################################
            itjTNW = 0
            itjNDW = 0
            for i in self.interjectList2:
                itjTNW += self.interjectList2[i]
                itjNDW += 1
            if itjTNW is 0:
                txt.write('\t감탄사 TNW : ' + str(0) + '\n')
                txt.write('\t감탄사 NDW : ' + str(0) + '\n')
                txt.write('\t감탄사 TTR : ' + str(0) + '\n' + '\n')
            else:
                txt.write('\t감탄사 TNW : ' + str(round(itjTNW, 2)) + '\n')
                txt.write('\t감탄사 NDW : ' + str(round(itjNDW, 2)) + '\n')
                txt.write('\t감탄사 TTR : ' + str(round((itjTNW / itjNDW), 2)) + '\n' + '\n')
            #######################################

            #######################################
            josaTNW = 0
            josaNDW = 0
            for i in self.josaList2:
                josaTNW += self.josaList2[i]
                josaNDW += 1
            if josaTNW is 0:
                txt.write('\t조사 TNW : ' + str(0) + '\n')
                txt.write('\t조사 NDW : ' + str(0) + '\n')
                txt.write('\t조사 TTR : ' + str(0) + '\n' + '\n')
            else:
                txt.write('\t조사 TNW : ' + str(round(josaTNW, 2)) + '\n')
                txt.write('\t조사 NDW : ' + str(round(josaNDW, 2)) + '\n')
                txt.write('\t조사 TTR : ' + str(round((josaTNW / josaNDW), 2)) + '\n' + '\n')
            #######################################

            #######################################
            eomiTNW = 0
            eomiNDW = 0
            for i in self.eomiList2:
                eomiTNW += self.eomiList2[i]
                eomiNDW += 1
            if eomiTNW is 0:
                txt.write('\t어미 TNW : ' + str(0) + '\n')
                txt.write('\t어미 NDW : ' + str(0) + '\n')
                txt.write('\t어미 TTR : ' + str(0) + '\n' + '\n')
            else:
                txt.write('\t어미 TNW : ' + str(round(eomiTNW, 2)) + '\n')
                txt.write('\t어미 NDW : ' + str(round(eomiNDW, 2)) + '\n')
                txt.write('\t어미 TTR : ' + str(round((eomiTNW / eomiNDW), 2)) + '\n' + '\n')

                #######################################





















