#Hallmark Christmas movie bingo card generator. Merry X-Mas.

import random
import numpy as np
import xlsxwriter

def generateBoard(boardnum,outfile):
    all_options=["mistletoe","a competition", "elderly person says something wise","city scene during opening credits",
                 "best friend makes mean comment to protagonist","son/daughter who wants their single mom/dad to get married",
                 "Christmas song as background music","a character who recounts a story from their childhood","sipping hot chocolate",
                 "corporate/business setting","ice skating","interrupted first kiss","flirtatious snowball fight","scene of decorating for Christmas/putting up lights",
                 "a handwritten note or letter","wedding dress","pretend to be my girlfriend/boyfriend/fiance",
                 "purchasing/wrapping/opening gifts","main characters meet and don't like each other at first",
                 "holiday baking/flour fight","airplane/airport","Santa","accidental fall that requires help up",
                 "Christmas/holiday party","a not-so-unexpected twist","someone lies badly","snowing","ugly Christmas sweater",
                 "carolers in old-fashioned costumes","Christmas festival","breakup","wearing open coat outside","candles",
                 "ex tries to win protagonist back","opening front door without a key","dead parent","mistaken identity", "snow on green grass",
                 "doting elderly relative","hanging out in local coffee shop or diner","non-traditional Christmas carol","someone mentions the magic of Christmas"]

    workbook = xlsxwriter.Workbook(outfile)

    for x in range(0,boardnum):
        worksheet = workbook.add_worksheet()
        bingoset=np.random.choice(all_options,24,replace=False)

        def add_to_board(board_range_low,board_range_high,row):
            while board_range_low<board_range_high:
                col=0
                r=bingoset[board_range_low:board_range_low+5]
                for i in range(len(r)):
                    worksheet.write(row, col+i,r[i])
                row+=1
                board_range_low+=5
        def middle_row(lst_start,lst_end,row,col):
            r=bingoset[lst_start:lst_end]
            for i in range(len(r)):
                worksheet.write(row,col+i,r[i])
        #first 2 rows
        add_to_board(0,10,0)
      
        #middle
        middle_row(10,12,2,0)
        worksheet.write(2,2,"FREE")
        middle_row(12,14,2,3)
        
        #last 2 rows
        add_to_board(14,24,3)
    workbook.close()
 
generateBoard(8,"xmas2.xlsx")
