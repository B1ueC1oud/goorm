from openpyxl import load_workbook
import os

load_wb = load_workbook("./quizdata.xlsx", data_only=True)
load_ws = load_wb['First Sheet']

for i in range(2,94):
    name='F'+str(i)
    print(name)
    print(load_ws[name].value)
    dirname=load_ws[name].value
    if '/' in dirname:
        dirname=dirname.replace("/", "_")
    output_dir = "./programming/"+dirname
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    # elif os.path.exists(output_dir):
    #     continue
    code = 'AA' + str(i)
    if load_ws[code].value is not None:
        print(output_dir + '/' + dirname + '_뼈대코드.txt')
        f = open(output_dir + '/' + dirname + '_뼈대코드.txt', 'w')
        f.write(load_ws[code].value)
        f.close()
    else:
        print("엑셀에서 "+"뼈대코드 "+code+" 비어있습니다.")
    prob = 'G' + str(i)
    if load_ws[prob].value is not None:
        f = open(output_dir + '/' + dirname + '_문제.html', 'w')
        f.write("<p>" + dirname + "</p>")
        f.write('<br>')
        f.write('<br>')
        f.write(load_ws[prob].value)
        f.close()

    else:
        print("엑셀에서 "+"문제html "+prob+" 비어있습니다.")

    testcase_num = 'H' + str(i)
    input_test='J'+ str(i)
    output_test='K'+ str(i)
    if load_ws[input_test].value is not None:
        f = open(output_dir + '/' + dirname + '_테스트입력.txt', 'w')
        f.write('테스트개수: ' + load_ws[testcase_num].value)
        f.write('\n')
        f.write('\n')
        input_split = load_ws[input_test].value
        f.write("\n=====================================\n")
        input_split = input_split.replace("::", "\n=====================================\n")
        f.write(input_split)
        f.close()
    else:
        print("엑셀에서 "+"테스트입력 "+input_test+" 비어있습니다.")
    if load_ws[output_test].value is not None:
        f = open(output_dir + '/' + dirname + '_테스트출력.txt', 'w')
        f.write('테스트개수: ' + load_ws[testcase_num].value)
        f.write('\n')
        f.write('\n')
        output_split = load_ws[output_test].value
        f.write("\n=====================================\n")
        output_split = output_split.replace("::", "\n=====================================\n")
        f.write(output_split)
        f.close()
    else:
        print("엑셀에서 "+"테스트출력 "+output_test+" 비어있습니다.")

