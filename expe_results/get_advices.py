from pyexcel_ods import get_data

class Person:
    def __init__(self, name, group, advices=None):
        self.name = name
        self.group = group
        self.advices = advices
        if not self.advices:
            self.advices = []



def get_all_advices_given(path, advice_dict):
    print('Getting data')
    data = get_data(path)
    print('Data aquired')

    for name, sheet in data.items():
        try:
            int(name)
        except:
            continue

        print(name)

        group = sheet[3][0]
        if 'G' not in str(group):
            group = 'G' + str(group)

        person = Person(name, group)
        # 3 because the last one isn't important
        advices = [sheet[12 + x*10][1].replace('/', '').replace('  ', ' ').lower() for x in range(3)]
        person.advices = advices
        advice_dict[group].append(person)

    return advice_dict


advices_given = {'G1': [], 'G2': [], 'G3': []}
# path = r'C:\Users\quentin\Documents\These\Databases\Darts\students_2\results_2_without_lefthanded.ods'
# advices_given = get_all_advices_given(path, advices_given)
path = r'Precision All data.ods'
advices_given = get_all_advices_given(path, advices_given)

number_advices = {'elbow move': {'j12': 0, 'j23': 0, 'j34': 0},
                  'align arm': {'j12': 0, 'j23': 0, 'j34': 0},
                  'javelin': {'j12': 0, 'j23': 0, 'j34': 0},
                  'leaning': {'j12': 0, 'j23': 0, 'j34': 0}}

for v in advices_given.values():
    for person in v:
        for i, advice in enumerate(person.advices):
            if 'elbow' in advice:
                number_advices['elbow move']['j'+str(i+1)+str(i+2)] += 1
            if 'align' in advice:
                number_advices['align arm']['j'+str(i+1)+str(i+2)] += 1
            if 'javelin' in advice:
                number_advices['javelin']['j'+str(i+1)+str(i+2)] += 1
            if 'leaning' in advice:
                number_advices['leaning']['j'+str(i+1)+str(i+2)] += 1

number_advices = {'G1': {'j12': {'elbow move': 0, 'align arm': 0, 'javelin': 0, 'leaning': 0},
                         'j23': {'elbow move': 0, 'align arm': 0, 'javelin': 0, 'leaning': 0},
                         'j34': {'elbow move': 0, 'align arm': 0, 'javelin': 0, 'leaning': 0}},
                  'G2': {'j12': {'elbow move': 0, 'align arm': 0, 'javelin': 0, 'leaning': 0},
                         'j23': {'elbow move': 0, 'align arm': 0, 'javelin': 0, 'leaning': 0},
                         'j34': {'elbow move': 0, 'align arm': 0, 'javelin': 0, 'leaning': 0}},
                  'G3': {'j12': {'elbow move': 0, 'align arm': 0, 'javelin': 0, 'leaning': 0},
                         'j23': {'elbow move': 0, 'align arm': 0, 'javelin': 0, 'leaning': 0},
                         'j34': {'elbow move': 0, 'align arm': 0, 'javelin': 0, 'leaning': 0}}}

# k = group
for k, v in advices_given.items():
    # v = person
    for person in v:
        # i = jeu
        for i, advice in enumerate(person.advices):
            if 'elbow' in advice:
                number_advices[k]['j'+str(i+1)+str(i+2)]['elbow move'] += 1
            if 'align' in advice:
                number_advices[k]['j'+str(i+1)+str(i+2)]['align arm'] += 1
            if 'javelin' in advice:
                number_advices[k]['j'+str(i+1)+str(i+2)]['javelin'] += 1
            if 'leaning' in advice:
                number_advices[k]['j'+str(i+1)+str(i+2)]['leaning'] += 1
breakpoint()
