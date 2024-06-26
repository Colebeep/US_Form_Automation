from io import StringIO
from html.parser import HTMLParser
#Solution taken from stack overflow:
#   https://stackoverflow.com/questions/753052/strip-html-from-strings-in-python
#   Answer by a user called Eloff, edited by a user called Olivier Le Floch
#   use to strip the html tags that are in the description object
class MLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.reset()
        self.strict = False
        self.convert_charrefs= True
        self.text = StringIO()
    def handle_data(self, d):
        self.text.write(d)
    def get_data(self):
        return self.text.getvalue()

#does the stripping, also calls helper function to get rid of the extra newlines that the tags needed, fixes some arrent spacing as well
def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return remove_extra_new_lines(s.get_data()).replace('    ',' ')

#must be an easier way to do this (regex)
#find a new line not after ending punc and remove it
#  needed as the form adds extra newlines to fit and looks bad in form
#  will fail to remove if the line naturally ends in punc, but that is uncommon occurence
#   could look for two newlines and remove all the ones that aren't in pairs perhaps?
def remove_extra_new_lines(s):
    ns = ''
    for i,x in enumerate(s):
        if x == '\n' and (s[i-1] not in '.?!'):
            continue
        else:
            ns += x
    return ns.replace('\n\n','\n')

from mailmerge import MailMerge #merges content into a docx template
import pandas as pd             #used to load excel files
import os                       #used for path exists and makedr

#not used in code yet, but is used as reference
Area_names ={
    'ca': 'Civitate A',
    'cb': 'Civitate B',
    't': 'Tesoro'
    }
    #path to template file
template = "./US Form Template.docx"
    #path to file holding trench coordinates for the season, will need to be updated each year
Trench_Coordinates = './Trench_Coordinates.csv'

#excel file is one workbook with multiple pages, this is the best way I found to read them all
#   x is a dictionary of dataframes
xls = pd.ExcelFile('kobo_input_all.xlsx')
x = pd.read_excel(xls, sheet_name=None)

#coord csv has a column for each trench, the header is the short name (Civitate A trench 95 becomes CA95)
#   the rows are written starting with NW going clockwise
#   coor_d{} is a dictionary lookup to speed up making multiple locus forms for the same trench
coorddf = pd.read_csv(Trench_Coordinates)
coor_d = {}

#load the main form sheet from the dict x, and replace all empty values with an empty string to legalize concat
summarydf = x['Locus Summary Entry 2023']
summarydf = summarydf.fillna('')

    #for every locus that has a filled form
for index, row in summarydf.iterrows():
    
    #load the template at path and to populate at end
    document = MailMerge(template)

    #trench ids are saved like this in sheet, so break apart relevant info
    # {shortened}{Trench_numer}_{year of excavation}
    id_break = row['Trench_ID'].split('_')

    #sanity check, year of trench id matches year of current season
    assert(id_break[1] == str(row['Season']))

    #will need to be expanded to include other designations
    #could be made quicker/easier with the dictionary of names
    if id_break[0].startswith('t') :
        area = "Tesoro"
        short = "T"
        trench_number = id_break[0][1:]
    elif id_break[0].startswith('ca'):
        area = "Civitate A"
        short = "CA"
        trench_number = id_break[0][2:]
    elif id_break[0].startswith('cb'):
        area = "Civitate B"
        short = "CB"
        trench_number = id_break[0][2:]
    else:
        area = "NOT YET IMPLEMENTED"
        short = "NA"
        trench_number = "ERROR"

        #id is based on the order the forms are done, and are only used to relate other sheets 
            #entries back with the main sheet
    current_id = row['_index']
    #create paths to save file and folder structure
    t_name = short + trench_number
    output_path = "./Output/" + t_name +'/'
    output_name = "US Locus " + str(row["Locus_ID"]) + '.docx'

    #if the shortened name isn't in the coordinate dictionary, add it will stripping out useless stuff
    if t_name not in coor_d:
        coor_d[t_name] = '\n'.join(coorddf[t_name].dropna().tolist())
    coor = coor_d[t_name]

    #check if folder for trench has been created yet
    if not os.path.exists(output_path):
        os.makedirs(output_path)
    else:
        #check to see if this exact form has been made before
        if os.path.exists(output_path+output_name) :
            #don't redo any US forms you have done before
            #skip to next
            #pass
            continue
        
    #all sections are used to grab data from excel and reformat it for the US Form

    #natural/artificial check
    if row['Natural_or_Artificial'] == 'natural':
        nat = 'X'
        art = ''
    else:
        nat = ''
        art = 'X'

    #plan numbers
    plans = ''
    try:
        for i,r in x['begin_repeat_Et0mtvMid'].iterrows():
            if r['_parent_index'] == current_id:
                if plans != '':
                    plans += '\n '
                plans += r['Associated_Plans']
    except Exception as e:
        print("-1",e)

    #grab all photo names
    photos = ''
    try:
        for i,r in x['begin_repeat_NawXUcMgh'].iterrows():
            if r['_parent_index'] == current_id:
                if photos != '':
                    photos += '\n'
                photos += r['Associated_Photos']
    except Exception as e:
        print("-3",e)

    #grab all special artifacts numbers, just need to be the number ie PC20230066 is just 66 
        #or sf-t103-2023-4-8 becomes 8
    specials = ''
    try:
        for i,r in x['begin_repeat_pyQUfbbMJ'].iterrows():
            if r['_parent_index'] == current_id:
                if specials != '':
                    specials += ', '
                if '-' in r['Contains_Special_Finds']:
                    lis = r['Contains_Special_Finds'].split('-')
                    specials += lis[-1]
                else:
                    specials += r['Contains_Special_Finds']
    except Exception as e:
        print("-2",e)
    if specials == '':
        specials = 'None'

    #tablulated counts structured {type}: {quantity} {measurement}\n  
    #                         ie: Tile: 13 Bowls\n
    tab = ''
    try:
        for i,r in x['begin_repeat_It8bliv1T'].iterrows():
            if r['_parent_index'] == current_id:
                new_str = ''
                if tab != '':
                    new_str += '\n'
                if r['Tabulated_Material_Type'] == 'Other':
                    new_str += str(r['Tabulated_Material_Type_Other'])
                else:
                    new_str += str(r['Tabulated_Material_Type'])
                new_str += ': ' + str(r['Tabulated_Count']) + ' ' + str(r['Tabulated_Count_Type'])
                if(not new_str.startswith('nan: nan') and ': 0.0 ' not in new_str):
                    if('Ceramic' in new_str ):
                        if(': 1.0 ' in new_str):
                            new_str = new_str.replace('Objects','Sherd')
                        else:
                            new_str = new_str.replace('Objects','Sherds')
                    elif('Objects' in new_str):
                        if(': 1.0 ' in new_str):
                            new_str = new_str.replace('Objects','Fragment')
                        else:
                            new_str = new_str.replace('Objects','Fragments')
                    if('.0 ' in new_str):
                        new_str = new_str.replace('.0 ', ' ')
                    tab += new_str.replace('_',' ')
    except Exception as e:
        print("-2",e)
    if tab=='':
        tab='None'

    #datable elements, simply grab whatever they input in the form comma sep
    datable = ''
    try:
        for i,r in x['begin_repeat_UlorltmZm'].iterrows():
            if r['_parent_index'] == current_id:
                if datable != '':
                    datable += ', '
                datable += str(r['Dateable_Elements'])
    except Exception as e:
        print("-1",e)
    if datable =='':
        datable = 'None'
        
        
    #period refers to the phase name (orientalizing, archaic)
    period = ', '.join(row['Locus_Period'].split(' ')).replace('_', ' ').capitalize()
    #prelim phasing refers to the year range
    prelim = row['Preliminary_Phasing'].replace('_', ' to ').capitalize()

    #Positioning/stratig
    #   all are the same structure comma seperated values
    #   try-excepts are used to as if in a season, no locus have a specific stratig type it will not appear in dicitonary, leading to a key error
    #       once one locus has one of that type the key error will stop
    #           if none have an anterior to locus code shouldn't crash
    #for ant/post
    ant = ''
    post = ''
    try:
        for i,r in x['group_strat_Anterior_To'].iterrows():
            if r['_parent_index'] == current_id:
                if ant != '':
                    ant += ', '
                ant += str(r['Anterior_To_Locus'])
    except:
        print('uh oh')
        pass
    if(ant != ''):
        if(', ' in ant):
            ant = 'Loci: ' + ant
        else:
            ant = 'Locus ' + ant 

    try:
        for i,r in x['group_strat_Posterior_To'].iterrows():
            if r['_parent_index'] == current_id:
                if post != '':
                    post += ', '
                post += str(r['Posterior_To_Locus'])
    except Exception as e:
        print(e)
        print('uh oh')
        pass
    if(post != ''):
        if(', ' in post):
            post = 'Loci: ' + post
        else:
            post = 'Locus ' + post 

    #10 stratig blocks in the template form
    above,overlies,underlies,is_cut_by,cuts,same_as,is_bound_to,fills,is_filled_by,below = '','','','','','','','','',''
    #above goes to SI APPOGGIA A ???
    try:
        for i,r in x['group_strat_Above'].iterrows():
            if r['_parent_index'] == current_id:
                if(above != ''):
                    above += ', '
                else:
                    if ',' in str(r['Above_Locus']) or ' ' in str(r['Above_Locus']):
                        above = 'Loci: '
                    else:
                        above = 'Locus '
                above += str(r['Above_Locus'])
    except Exception as p:
        print(p, 0)
        pass
    if(',' in above and 'Locus' in above):
        above = above.replace('Locus', 'Loci:')
    #Overlies goes to COPRE
    try:
        for i,r in x['group_strat_Overlies'].iterrows():
            if r['_parent_index'] == current_id:
                if(overlies != ''):
                    overlies += ', '
                else:
                    if ',' in str(r['Overlies_Locus']) or ' ' in str(r['Overlies_Locus']):
                        overlies = 'Loci: '
                    else:
                        overlies = 'Locus '
                overlies += str(r['Overlies_Locus'])
    except Exception as p:
        print(p, 1)
        pass
    if(',' in overlies and 'Locus' in overlies):
        overlies = overlies.replace('Locus', 'Loci:')

    #Underlies goes to COPERTO DA
    try:
        for i,r in x['group_strat_Underlies'].iterrows():
            if r['_parent_index'] == current_id:
                if(underlies != ''):
                    underlies += ', '
                else:
                    if ',' in str(r['Underlies_Locus']) or ' ' in str(r['Underlies_Locus']):
                        underlies = 'Loci: '
                    else:
                        underlies = 'Locus '
                underlies += str(r['Underlies_Locus'])
    except Exception as p:
        print(p, 2)
        pass
    if(',' in underlies and 'Locus' in underlies):
        underlies = underlies.replace('Locus', 'Loci:')

    #Is Cut By goes to TAGLIATO DA
    try:
        for i,r in x['group_strat_Is_Cut_By'].iterrows():
            if r['_parent_index'] == current_id:
                if(is_cut_by != ''):
                    is_cut_by += ', '
                else:
                    if ',' in str(r['Is_Cut_By_Locus']) or ' ' in str(r['Is_Cut_By_Locus']):
                        is_cut_by = 'Loci: '
                    else:
                        is_cut_by = 'Locus '
                is_cut_by += str(r['Is_Cut_By_Locus'])
    except Exception as p:
        print(p, 3)
        pass
    if(',' in is_cut_by and 'Locus' in is_cut_by):
        is_cut_by = is_cut_by.replace('Locus', 'Loci:')

    #Cuts goes to TAGLIA
    try:
        for i,r in x['group_strat_Cuts'].iterrows():
            if r['_parent_index'] == current_id:
                if(cuts != ''):
                    cuts += ', '
                else:
                    if ',' in str(r['Cuts_Locus']) or ' ' in str(r['Cuts_Locus']):
                        cuts = 'Loci: '
                    else:
                        cuts = 'Locus '
                cuts += str(r['Cuts_Locus'])
    except Exception as p:
        print(p, 4)
        pass
    if(',' in cuts and 'Locus' in cuts):
        cuts = cuts.replace('Locus', 'Loci:')

    #Same as goes to UGUALE A
    try:
        for i,r in x['group_strat_Same_As'].iterrows():
            if r['_parent_index'] == current_id:
                if(same_as != ''):
                    same_as += ', '
                else:
                    if ',' in str(r['Same_As_Locus']) or ' ' in str(r['Same_As_Locus']):
                        same_as = 'Loci: '
                    else:
                        same_as = 'Locus '
                same_as += str(r['Same_As_Locus'])
    except Exception as p:
        print(p, 5)
        pass
    if(',' in same_as and 'Locus' in same_as):
        same_as = same_as.replace('Locus', 'Loci:')

    #Is Bound To goes to SI LEGA A
    try:
        for i,r in x['group_strat_Is_Bound_To'].iterrows():
            if r['_parent_index'] == current_id:
                if(is_bound_to != ''):
                    is_bound_to += ', '
                else:
                    if ',' in str(r['Is_Bound_To_Locus']) or ' ' in str(r['Is_Bound_To_Locus']):
                        is_bound_to = 'Loci: '
                    else:
                        is_bound_to = 'Locus '
                is_bound_to += str(r['Is_Bound_To_Locus'])
    except Exception as p:
        print(p, 6)
        pass
    if(',' in is_bound_to and 'Locus' in is_bound_to):
        is_bound_to = is_bound_to.replace('Locus', 'Loci')

    #Below goes to ????
    try:
        for i,r in x['group_strat_Below'].iterrows():
            if r['_parent_index'] == current_id:
                if(below != ''):
                    below += ', '
                else:
                    if ',' in str(r['Below_Locus']) or ' ' in str(r['Below_Locus']):
                        below = 'Loci: '
                    else:
                        below = 'Locus '
                below += str(r['Below_Locus'])
    except Exception as p:
        print(p, 7)
        pass
    if(',' in below and 'Locus' in below):
        below = below.replace('Locus', 'Loci:')

    #Fills goes to RIEMPIE
    try:
        for i,r in x['group_strat_Fills'].iterrows():
            if r['_parent_index'] == current_id:
                if(fills != ''):
                    filss += ', '
                else:
                    if ',' in str(r['Fills_Locus']) or ' ' in str(r['Fills_Locus']):
                        fills = 'Loci: '
                    else:
                        fills = 'Locus '
                fills += str(r['Fills_Locus'])
    except Exception as p:
        print(p, 8)
        pass
    if(',' in fills and 'Locus' in fills):
        fills = fills.replace('Locus', 'Loci:')

    #Is Filled By goes to RIEMPITO DA
    try:
        for i,r in x['group_strat_Is_Filled_By'].iterrows():
            if r['_parent_index'] == current_id:
                if(is_filled_by != ''):
                    is_filled_by += ', '
                else:
                    if ',' in str(r['Is_Filled_By_Locus']) or ' ' in str(r['Is_Filled_By_Locus']):
                        is_filled_by = 'Loci: '
                    else:
                        is_filled_by = 'Locus '
                is_filled_by += str(r['Is_Filled_By_Locus'])
    except Exception as p:
        print(p, 9)
        pass
    if(',' in is_filled_by and 'Locus' in is_filled_by):
        is_filled_by = is_filled_by.replace('Locus', 'Loci:')

    #Recovery techniques are lacking in detail in the Kobo form, so similarly the detail is low in US form
    #should be fine, a simple yes/no for the recovery should be okay
    siev = ''
    if str(row['Recovery_Techniques/Dry_Sieved']) == '1.0':
        siev += 'Dry Sieved'
    if str(row['Recovery_Techniques/Wet_Sieved']) == '1.0':
        if siev != '':
            siev += ' and '
        siev += "Wet Sieved"
    if siev == '':
        siev = 'None'
    flotation = ''
    if str(row['Recovery_Techniques/Flotation']) == '1.0':
        flotation += 'Floatation Samples Taken'
    if flotation == '':
        flotation = 'None'

    #list of what samples were taken
    samples = ''
    try:
        for i,r in x['begin_repeat_v62FStDQL'].iterrows():
            if r['_parent_index'] == current_id:
                if samples != '':
                    samples +='\n'
                samples += str(r['Contains_Samples'])
    except Exception as p:
        print(p, 10)
        pass
    if samples =='':
        samples = 'None'

        #complex data has been grabbed, the rest can be done in line
    document.merge(
        Locus_Number = str(row['Locus_ID']),
        Locus_Number_2 = str(row['Locus_ID']),
        Year = str(row['Season']),
        Area = area,
        Trench_Number = trench_number,
        Trench_Coordinates = coor,
        Elevation_or_Depth = str(row['Depth_cm']) + ' cm deep',
        Natural = nat,
        Artificial = art,
        Plan_Number = plans,
        Photo_Names = photos,
        Catalogue_Numbers = specials,
        Definition_and_Position = str(row['Definition_and_Position']),
        Criteria_for_Distinction = str(row['Criteria_for_Distinction']),
        Mode_of_Formation = str(row['Mode_of_Formation']),
        Inorganic_List = str(row['Inorganic_Components']),
        Organic_List = str(row['Organic_Components']),
        Consistency = (str(row['Deposit_Compaction']) + ' compaction').capitalize(),
        Color = str(row['Munsell_Color']),
        Dimensions = str(row['Width_m']) + ' m wide x ' + str(row['Length_m']) + ' m long',
        State_of_Conservation = str(row['State_of_Conservation']),
        Description = strip_tags(str(row['Description'])),
        Equal_To = same_as,          #UGUALE A
        Is_Bound_To = is_bound_to,   #SI LEGA A
        Leans_On = overlies,            #GLI SI APPOGGIA (should go to covers)
        Is_Above = underlies,            #SI APPOGGIA A  (should go to covered by)
        Covered_By = below,      #COPERTO DA
        Covers = above,           #COPRE
        Is_Cut_By = is_cut_by,       #TAGLIATO DA
        Cuts = cuts,                 #TAGLIA
        Is_Filled_By = is_filled_by, #RIEMPITO DA
        Fills = fills,               #RIEMPIE
        Previous_Loci = post,        #POSTERIORE A
        Following_Loci = ant,        #ANTERIORE A
        Observations = str(row['Observations']),
        Interpretations = str(row['Interpretation']),
        Datable_Elements = datable,
        Dating = prelim,
        Period_or_Phase = str(period),
        Quantity_Report = tab,
        Samples = samples,
        Flotation = flotation,
        Sieving = siev,
        Reliability = str(row['Stratigraphic_Reliability']).capitalize()
    )

    document.write(output_path + output_name)