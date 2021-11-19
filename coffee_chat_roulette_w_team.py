import pandas as pd
import random
import xlsxwriter


class record:
    '''
    Object containing a person's name, team affiliation, and historical pairings
    '''
    def __init__(self, name, team, hist, teamhist):
        self.name = name
        self.team = team
        self.hist = hist
        self.teamhist = teamhist
    def forget(self, mem):
        hist_list = self.hist.split(',')
        if len(hist_list) > mem:
            self.hist = ','.join(hist_list[-mem:])
        teamhist_list = self.teamhist.split(',')
        if len(teamhist_list) > mem:
            self.teamhist = ','.join(teamhist_list[-mem:])
    def show(self):
        print(self.name, self.team, self.hist)


def build_people(input_file):
    '''
    Builds a dictionary of <record> objects pulled from the input file 
    '''
    people = {};
    input_data = pd.read_excel(input_file, sheet_name = 'History')
    num_peeps  = len(input_data.index) 

    for ind in range(num_peeps):
        name = input_data.loc[ind,'Your Name']
        team = input_data.loc[ind,'Team Name']
        hist = input_data.loc[ind,'History']
        teamhist = input_data.loc[ind,'Team History']

        peep = record(name, team, hist, teamhist)
        people[peep.name] = peep

    return people

def assign_pairs(names, people,  team_hist_save = 100000,
                 debug = False):
    '''
    Assigns pairs of people from a list of names, 
    checking for cross team pairing and no history
    violations, based on a people library
    '''
    pairs = []
    pairing_good = False;
    num_iterations = 0
    relax = False
    while not pairing_good:
        num_iterations += 1;
        if num_iterations % team_hist_save == 0 and not relax: 
            check = input("It's been %ik iterations. Relax Team History Constraint? (y/n) "%(num_iterations/1000))
            if check == 'y':
                relax = True
                
            
        pairing_good = True;
        random.shuffle(names)
        pairs = [[names[k], names[k+1]] for k in range(0,len(names)-1, 2)]
        for p in pairs:
            lead = people[p[0]]
            foll = people[p[1]]

            if lead.team == foll.team:
                pairing_good = False; 
                if debug: print('Same Team', lead.name, foll.name)
                break 
            if foll.name in lead.hist:
                pairing_good = False; 
                if debug: print('Recent Team',lead.name, foll.name)
                break
            if lead.name in foll.hist:
                pairing_good = False; 
                if debug: print('Recent Team',lead.name, foll.name)
                break
            if not relax:
                if foll.team in lead.teamhist:
                    pairing_good = False; 
                    if debug: print('Recent Team',lead.name, foll.name)
                    break
                if lead.team in foll.teamhist:
                    pairing_good = False; 
                    if debug: print('Recent Team',lead.name, foll.name)
                    break
            if 'OR - ' in lead.team and 'OR - ' in foll.team:
                pairing_good = False;
                if debug: print('Two OR people...', lead.name, foll.name)
                break
                
            else:
                if foll.team == lead.teamhist.split(',')[-1]:
                    pairing_good = False; 
                    if debug: print('Recent Team',lead.name, foll.name)
                    break
                if lead.team == foll.teamhist.split(',')[-1]:
                    pairing_good = False; 
                    if debug: print('Recent Team',lead.name, foll.name)
                    break                


    return(pairs)

def update_history(pairs, people, memory = 2):
    '''
    Updates the history records of people based on current pairs

    <memory> object defines number of names to maintain in history
    '''
    for p in pairs:
        for x,y in zip([0,1],[1,0]):
            people[p[x]].hist += ',' + people[p[y]].name
            people[p[x]].teamhist += ',' + people[p[y]].team
            people[p[x]].forget(memory)
    return()

if __name__ == '__main__':

    input_file = 'ccr_input_2004.xlsx'   #Name of input file
    output_file = 'ccr_output_2004.xlsx' #Name of output file (make the same to overwrite)
    forced_double = 'Eamonn Shirey'      #Name of person who doubles up when odd # of participants
    debug = False
    memory = 2

    people = build_people(input_file)  #Build dictionary of people
    names = list(people.keys())        #Pull names of people
    
    
    # Assign someone two pairs if you have an odd number of people
    if len(names) % 2 == 1: 
        
        if forced_double != None: double_duty = forced_double
        else: double_duty = random.choice(names)
        names.append(double_duty)
        print('%s is pulling double duty'%double_duty)


    pairs = assign_pairs(names, people, debug = debug) #Assign the pairs
    update_history(pairs,people, memory = 2) #Update history with new pairs


    #Print out and save the results to a file
    print('------')
    for p in pairs: print('%s, %s\t | \t\t %s, %s'%(p[0],people[p[0]].team,p[1], people[p[1]].team))



    workbook = xlsxwriter.Workbook(output_file) 
    worksheet = workbook.add_worksheet('History');

    worksheet.write(0, 0, 'Your Name')
    worksheet.write(0, 1, 'Team Name')
    worksheet.write(0, 2, 'History')
    worksheet.write(0, 3, 'Team History')

    row = 1;
    for p in people:
        worksheet.write(row, 0, people[p].name)
        worksheet.write(row, 1, people[p].team)
        worksheet.write(row, 2, people[p].hist)
        worksheet.write(row, 3, people[p].teamhist)
        row += 1


    results = workbook.add_worksheet('Results');
    row = 1;
    for p in pairs:
        results.write(row, 0, p[0])
        results.write(row, 1, people[p[0]].team)
        results.write(row, 2, p[1])
        results.write(row, 3, people[p[1]].team)
        row += 1
    workbook.close()






