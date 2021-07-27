import pickle
import glob
import csv

def save_obj(obj, name):
    with open(name + '.pkl', 'wb') as f:
        pickle.dump(obj, f, pickle.HIGHEST_PROTOCOL)

def load_obj(name):
    with open(name + '.pkl', 'rb') as f:
        return pickle.load(f)

def mergeTextIntoCSV(dir):
    dic = load_obj('abreviationdict')

    txtfiles = glob.glob(dir + '/*.txt')
    with open(dir + '/mergedfiles.csv', 'w', newline='', encoding='utf8') as csvfile:
        header = list(dic.values())
        header.append('Firm Name')
        writer = csv.DictWriter(csvfile, fieldnames=header)
        writer.writeheader()

        for file in txtfiles:
            row = {}
            prevAbr = ''
            with open(file, encoding='utf8') as FileObj:
                for lines in FileObj:
                    row['Firm Name'] = file.split('\\')[-1]
                    abrev = lines[0 : 2].replace('\n', '')
                    content = lines[3:].replace('\n', '')
                    if len(lines.replace('\n', '')) == 0:
                        writer.writerow(row)
                        row = {}
                        continue
                    if abrev in dic.keys():
                        row[dic[abrev]] = row.get(dic[abrev], '') + content
                        prevAbr = abrev
                    if abrev == '  ':
                        row[dic[prevAbr]] = row.get(dic[prevAbr], '') + ' ' + content
                print('Done converting and merging ' + file.split('\\')[-1] + ' into CSV')

dir = input('Enter location of .txt files to be merged into CSV: ')
mergeTextIntoCSV(dir)