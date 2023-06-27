for i in range(0,1):
    print("TRY : ",i)
    import pandas as pd
    import re as re
    regular_list = []

    def whitespace_remover(dataframe):
        # iterating over the columns
        for i in dataframe.columns:
            initialType = dataframe[i].dtype
            # checking datatype of each columns
            dataframe[i] = dataframe[i].astype(str)
            dataframe[i] = dataframe[i].str.strip()
            dataframe[i] = dataframe[i].astype(initialType)
        return dataframe
    
    def excelAppender(dictUniqueA1ID, dictToAppend, dfToAppend, colName):
        tempLst = []
        for k1, v1 in dictUniqueA1ID.items():
            for k2, v2 in dictToAppend.items():
                if k1 == k2:
                    tempLst.append(v2)
        dfToAppend[colName] = tempLst

    def dictIntersector(dict1, dict2):
        for k1, v1 in dict1.items():
            for k2, v2 in dict2.items():
                if k1 == k2:
                    dict1.update({k1: v2})
        return dict1

    def flatten(list_of_lists):
        if len(list_of_lists) == 0:
            return list_of_lists
        if isinstance(list_of_lists[0], type(list)):
            return flatten(list_of_lists[0]) + flatten(list_of_lists[1:])
        return list_of_lists[:1] + flatten(list_of_lists[1:])

    def val_append(dict_sub, key, value): #добавя нова стойност към съществуващ key
        if key in dict_sub:
            if not isinstance(dict_sub[key], type(list)):
                # converting key to list type
                dict_sub[key] = [dict_sub[key]]

                # Append the key's value in list
            dict_sub[key].append(value)
        return dict_sub

    def dictCreator(idVariable, attribute): #zip-ва един речник от два листа
        dict = {}
        for A, B in zip(idVariable, attribute):
            dict[A] = B
        return dict
    dismantledFilter = True
    #отваря всеки шийт в различен data frame
    cetinFileName = 'C:\\Users\\a1bg511027\\Desktop\\tablici\\NEW_Cetin tracking collocations.xlsx'
    vivaFileName = 'C:\\Users\\a1bg511027\\Desktop\\tablici\\Vivacom collocations.xlsx'
    dfCetinHostCetin = pd.read_excel(open(cetinFileName, 'rb'), sheet_name='On Air - Cetin host', index_col=None)
    dfCetinHostA1 = pd.read_excel(open(cetinFileName, 'rb'), sheet_name='On Air - A1 host', index_col=None)
    dfVivaHostViva = pd.read_excel(open(vivaFileName, 'rb'), sheet_name='On Air - Vivacom host', index_col=None)
    dfVivaHostA1 = pd.read_excel(open(vivaFileName, 'rb'), sheet_name='On Air - A1 host', index_col=None)
    vivacomRegexPattern = '\w{2}\d{4}'
    a1RegexPattern = '\w{3}\d{4}'
    dismantledFilter = False
    dfLst = [dfCetinHostCetin, dfCetinHostA1, dfVivaHostViva, dfVivaHostA1]
    for frame in dfLst:
        frame = whitespace_remover(frame)
    
    #намира абсолютно всички A1 SID от всички шийтове от таблиците и прави list само с уникалните такива
    cetinHostA1ID = dfCetinHostA1["A1 SID"].to_list()
    vivaHostA1ID = dfVivaHostA1["A1 SID"].to_list()

    guest1CetinHostCetin = dfCetinHostCetin["1's guest"].dropna().to_list()
    guest2CetinHostCetin = dfCetinHostCetin["2'nd guest"].dropna().to_list()
    guest1VivaHostViva = dfVivaHostViva["1's guest"].dropna().to_list()
    guest2VivaHostViva = dfVivaHostViva["2'nd guest"].dropna().to_list()
    guestVivaHostA1 = [guest1VivaHostViva, guest2VivaHostViva]
    guestLists = [guest1CetinHostCetin, guest2CetinHostCetin, guest1VivaHostViva, guest2VivaHostViva]
    guestListA1 = []
    for list in guestLists:
        for ele in list:
            ele = str(ele)
            if re.findall(a1RegexPattern, ele):
                guestListA1.append(ele)

    #на някои места има елементи от list по този начин \n['siteID']\n: цикъла ги премахва тези \n
    for ele in guestListA1:
        if re.findall("\n", ele):
            guestListA1.remove(ele)
            ele = ele.replace("\n", "")
            guestListA1.append(ele)

    setOfA1ID = set().union(guestListA1, cetinHostA1ID,vivaHostA1ID)

    dictUniqueA1ID = dict.fromkeys(setOfA1ID) #прави речник като има единствено ключове с уникалните ID-та на А1
    # после този dict ще се ъпдейтва с останалите речници с цел нагаждане на всеки сайт да се знае кои са участващите

    dictUniqueA1ID_Owner = dictUniqueA1ID.copy()
    dictUniqueA1ID_Sname = dictUniqueA1ID.copy()
    dictUniqueA1ID_Guest1 = dictUniqueA1ID.copy()
    dictUniqueA1ID_SiteOwner = dictUniqueA1ID.copy()

    #идеята е да се създаде dict на база A1 SID - Vivacom SID (guest);  A1 SID - Cetin SID
    cetinHostCetinID = dfCetinHostCetin["Cetin SID"].to_list()
    cetinHostCetinID = [str(x) for x in cetinHostCetinID]
    guest1CetinHostCetin = dfCetinHostCetin["1's guest"].to_list()
    guest2CetinHostCetin = dfCetinHostCetin["2'nd guest"].to_list()
    dictA1hostCetin = {}
    a1GuestIDhostCetin = []
    vivaGuestIDhostCetin = []
    row = 0
    for item in guest1CetinHostCetin:
        item = re.findall(a1RegexPattern, item)
        if len(item) == 0:
            item = guest2CetinHostCetin[row]
            a1GuestIDhostCetin.append(item)
        else:
            a1GuestIDhostCetin.append(item[0])
        row = row+1

    row = 0
    for item in guest2CetinHostCetin:
        item = str(item)
        if re.findall(a1RegexPattern,item):
            vivaGuestIDhostCetin.append(guest1CetinHostCetin[row])
        elif item == "nan":
            vivaGuestIDhostCetin.append(None)
        else: vivaGuestIDhostCetin.append(item)
        row = row+1
    dictA1GuestIDhostCetin_vivaGuestIDhostCetin = dictCreator(a1GuestIDhostCetin,vivaGuestIDhostCetin)
    dictA1GuestIDhostCetin_cetinHostCetinID = dictCreator(a1GuestIDhostCetin,cetinHostCetinID)

    #идеята е да се създаде dict на база A1 SID - Vivacom SID;  A1 SID - Cetin(guest) SID
    vivaHostVivaID = dfVivaHostViva["Vivacom SID"].to_list()
    guest1VivaHostViva = dfVivaHostViva["1's guest"].to_list()
    guest2VivaHostViva = dfVivaHostViva["2'nd guest"].to_list()
    dictA1hostViva = {}
    a1GuestIDhostViva = []
    cetinGuestIDhostViva = []

    row = 0
    for item in guest1VivaHostViva:
        item = str(item)
        if re.findall(a1RegexPattern,item):
            a1GuestIDhostViva.append(item)
        else:
            a1GuestIDhostViva.append(guest2VivaHostViva[row])
        row = row+1

    row = 0
    for item in guest2VivaHostViva:
        item = str(item)
        if re.findall(a1RegexPattern,item):
            item = str(guest1VivaHostViva[row])
            cetinGuestIDhostViva.append(item)
        elif item == "nan":
            cetinGuestIDhostViva.append(None)
        else:
            cetinGuestIDhostViva.append(item)
        row = row+1
    
    dictA1GuestIDhostViva_vivaIDhostViva = dictCreator(a1GuestIDhostViva,vivaHostVivaID)
    dictA1GuestIDhostViva_cetinGuestIDhostViva = dictCreator(a1GuestIDhostViva,cetinGuestIDhostViva)

    #идеята е да се създаде dict на база A1 SID(Host) - Vivacom SID(Guest) и A1 SID(Host) - CETIN SID(Guest) #file Vivacom Collocations
    vivaHostA1ID = dfVivaHostA1["A1 SID"].to_list()
    guest1VivaHostA1 = dfVivaHostA1["1's guest"].to_list()
    guest2VivaHostA1 = dfVivaHostA1["2'nd guest"].to_list()
    vivaFileVivaGuestIDhostA1 = []
    vivaFileCetinGuestIDhostA1= []
    row = 0
    for item in guest1VivaHostA1:
        item = str(item)
        if re.findall(vivacomRegexPattern,item):
            vivaFileVivaGuestIDhostA1.append(item)
        else: vivaFileVivaGuestIDhostA1.append(guest2VivaHostA1[row])
        row = row+1
    #**************ново************** 
    row = 0

    for item in guest2VivaHostA1:
        item = str(item)
        if re.findall(vivacomRegexPattern,item):
            item = str(guest1VivaHostA1[row])
            vivaFileCetinGuestIDhostA1.append(item)
        elif item == "nan":
            vivaFileCetinGuestIDhostA1.append(None)
        else:
            vivaFileCetinGuestIDhostA1.append(item)
        row = row+1

    
    dictFileVivaA1ID_guestCetinIDhostA1 = dictCreator(vivaHostA1ID, vivaFileCetinGuestIDhostA1)
    dictFileVivaA1ID_guestVivaIDhostA1 = dictCreator(vivaHostA1ID, vivaFileVivaGuestIDhostA1)

    #идеята е да се създаде dict на база A1 SID(Host) - CETIN SID(Guest) и A1 SID(Host) - Vivacom SID(Guest) #file Cetin Collocations
    cetinHostA1ID = dfCetinHostA1["A1 SID"].to_list()
    guest1CetinHostA1 = dfCetinHostA1["1's guest"].to_list()
    guest2CetinHostA1 = dfCetinHostA1["2'nd guest"].to_list()
    cetinFileVivaGuestIDhostA1 = []
    cetinFileCetinGuestIDhostA1= []
    row = 0
    for item in guest1CetinHostA1:
        item = str(item)
        if item.isnumeric():
            cetinFileCetinGuestIDhostA1.append(item)
        else:
            item = str(guest2CetinHostA1[row])
            cetinFileCetinGuestIDhostA1.append(item)
        row = row + 1

    row = 0
    for item in guest2CetinHostA1:
        item = str(item)
        if re.findall(vivacomRegexPattern,item):
            item = str(guest1CetinHostA1[row])
            cetinFileVivaGuestIDhostA1.append(item)
        elif item == "nan":
            cetinFileVivaGuestIDhostA1.append(None)
        else: cetinFileVivaGuestIDhostA1.append(item)
        '''
        if item == "Requested":
            item1 = guest1CetinHostA1[row]
            if item1.isnumeric():
                cetinFileVivaGuestIDhostA1.append(item)
            elif re.findall(vivacomRegexPattern,item1):
                cetinFileCetinGuestIDhostA1.append(item)
        else:
            item = re.findall(vivacomRegexPattern,item)
            if len(item) == 0:
                item = str(guest1CetinHostA1[row])
                if item.isnumeric():
                    cetinFileVivaGuestIDhostA1.append(None)
                else:
                    cetinFileVivaGuestIDhostA1.append(str(item))
            else:
                cetinFileVivaGuestIDhostA1.append(item[0])
                '''
        row = row + 1
    
    dictFileCetinA1ID_guestCetinIDhostA1 = dictCreator(cetinHostA1ID, cetinFileCetinGuestIDhostA1)
    dictFileCetinA1ID_guestVivaIDhostA1 = dictCreator(cetinHostA1ID, cetinFileVivaGuestIDhostA1)

    #dictA1GuestIDhostViva_vivaIDhostViva, dictA1GuestIDhostViva_cetinGuestIDhostViva ---- A1 Guest : Viva Host, A1 Guest : CETIN Guest - Viva Collocations
    #dictA1GuestIDhostCetin_vivaGuestIDhostCetin, dictA1GuestIDhostCetin_cetinHostCetinID ---- A1 Guest : Viva Guest, A1 Guest : CETIN Host - Cetin Collocations
    #dictFileVivaA1ID_guestCetinIDhostA1, dictFileVivaA1ID_guestVivaIDhostA1 ---- A1 Host: CETIN Guest, A1 Host: Viva Guest - Viva Collocations
    #dictFileCetinA1ID_guestCetinIDhostA1, dictFileCetinA1ID_guestVivaIDhostA1 ---- A1 Host: CETIN Guest, A1 Host: Viva Guest - Cetin Collocations

    dictList = [dictA1GuestIDhostViva_vivaIDhostViva.items(), dictA1GuestIDhostViva_cetinGuestIDhostViva.items(),
                   dictA1GuestIDhostCetin_vivaGuestIDhostCetin.items(), dictA1GuestIDhostCetin_cetinHostCetinID.items(),
                   dictFileVivaA1ID_guestCetinIDhostA1.items(), dictFileVivaA1ID_guestVivaIDhostA1.items(),
                   dictFileCetinA1ID_guestCetinIDhostA1.items(), dictFileCetinA1ID_guestVivaIDhostA1.items()]
    #създава общ речник от всичките с ключове на A1 SID
    for dict in dictList:
        for key, value in dict:
            val_append(dictUniqueA1ID, key, value)

    #маха повтарящи се ID-та и None от values
    for key, value in dictUniqueA1ID.items():
        value = flatten(value)
        filtered_list = []
        for ele in value:
            if ele != None and ele!='nan':
                filtered_list.append(ele)
        setTry = set(filtered_list)
        #if len(setTry)>2:
            #print(key, setTry)
        listAgain = []
        for i in setTry:
            listAgain.append(i)
        dictUniqueA1ID[key] = listAgain

    #=====================================================================================

    #===============Owners===============
    A1HostOwnerVivaFile = dfVivaHostA1["Owner"].to_list()
    A1HostOwnerCetinFile = dfCetinHostA1["Owner"].to_list()
    A1GuestOwnerVivaFile = dfVivaHostViva["Owner"].to_list()
    A1GuestOwnerCetinFile = dfCetinHostCetin["Owner"].to_list()
    dictA1ID_Owner_A1HostVivaFile = dictCreator(vivaHostA1ID, A1HostOwnerVivaFile)
    dictA1ID_Owner_A1HostCetinFile = dictCreator(cetinHostA1ID, A1HostOwnerCetinFile)
    dictA1ID_Owner_A1GuestVivaFile = dictCreator(a1GuestIDhostViva, A1GuestOwnerVivaFile)
    dictA1ID_Owner_A1GuestCetinFile = dictCreator(a1GuestIDhostCetin, A1GuestOwnerCetinFile)
    dictList = [dictA1ID_Owner_A1HostVivaFile, dictA1ID_Owner_A1HostCetinFile, dictA1ID_Owner_A1GuestVivaFile, dictA1ID_Owner_A1GuestCetinFile]
    for dict in dictList:
        dictUniqueA1ID_Owner = dictIntersector(dictUniqueA1ID_Owner, dict)
    #=================Owners End=================

    #=================Site Owner=================
    A1HostSiteOwnerVivaFile = len(vivaHostA1ID) * ['A1']
    A1HostSiteOwnerCetinFile = len(cetinHostA1ID) * ['A1']
    A1GuestSiteOwnerVivaFile = len(a1GuestIDhostViva) * ['Vivacom']
    A1GuestSiteOwnerCetinFile = len(a1GuestIDhostCetin) * ['CETIN']
    dictA1ID_SiteOwner_A1HostVivaFile = dictCreator(vivaHostA1ID, A1HostSiteOwnerVivaFile)
    dictA1ID_SiteOwner_A1HostCetinFile = dictCreator(cetinHostA1ID, A1HostSiteOwnerCetinFile)
    dictA1ID_SiteOwner_A1GuestVivaFile = dictCreator(a1GuestIDhostViva, A1GuestSiteOwnerVivaFile)
    dictA1ID_SiteOwner_A1GuestCetinFile = dictCreator(a1GuestIDhostCetin, A1GuestSiteOwnerCetinFile)
    dictList = [dictA1ID_SiteOwner_A1HostVivaFile, dictA1ID_SiteOwner_A1HostCetinFile, dictA1ID_SiteOwner_A1GuestVivaFile, dictA1ID_SiteOwner_A1GuestCetinFile]
    for dict in dictList:
        dictUniqueA1ID_SiteOwner = dictIntersector(dictUniqueA1ID_SiteOwner, dict)
    #===============Site Names===============
    a1HostA1SnameVivaFile = dfVivaHostA1["A1 Site Name"].to_list()
    a1HostA1SnameCetinFile = dfCetinHostA1["A1 Site Name"].to_list()
    cetinHostA1Sname= dfCetinHostCetin["A1 site name"].to_list()
    vivaHostA1Sname = dfVivaHostViva["A1 site name"].to_list()
    a1HostVivaSname = dfVivaHostA1['Vivacom site name'].to_list()
    vivaHostVivaSname = dfVivaHostViva['Vivacom Site Name'].to_list()
    a1HostCetinSname = dfCetinHostA1['CETIN site name'].to_list()
    cetinHostCetinSname = dfCetinHostCetin['CETIN Site Name'].to_list()

    #Dicts A1 ID Host/Guest : Site Name -> A1 ALL ID: Site Name
    dictA1ID_Sname_a1Host_vivaFile = dictCreator(vivaHostA1ID, a1HostA1SnameVivaFile)
    dictA1ID_Sname_a1Host_cetinFIle = dictCreator(cetinHostA1ID, a1HostA1SnameCetinFile)
    dictA1ID_Sname_a1guest_viva = dictCreator(a1GuestIDhostViva, vivaHostA1Sname)
    dictA1ID_Sname_a1guest_cetin = dictCreator(a1GuestIDhostCetin, cetinHostA1Sname)
    dictList = [dictA1ID_Sname_a1Host_vivaFile, dictA1ID_Sname_a1Host_cetinFIle, dictA1ID_Sname_a1guest_viva, dictA1ID_Sname_a1guest_cetin]
    for dict in dictList:
        dictUniqueA1ID_Sname = dictIntersector(dictUniqueA1ID_Sname, dict)

    #Dicts Viva ID Host/Guest : Site Name -> Viva ALL ID: Site Name
    dictVivaID_Sname = dictCreator(vivaHostVivaID, vivaHostVivaSname)
    dictVivaIDguest_Sname_a1 = dictCreator(vivaFileVivaGuestIDhostA1, a1HostVivaSname)
    dictVivaID_Sname.update(dictVivaIDguest_Sname_a1)

    #Dicts Cetin ID Host/Guest : Site Name -> Cetin ALL ID: Site Name
    dictCetinID_Sname = dictCreator(cetinHostCetinID, cetinHostCetinSname)
    dictCetinIDguest_Sname_a1 = dictCreator(cetinFileCetinGuestIDhostA1, a1HostCetinSname)
    dictCetinID_Sname.update(dictCetinIDguest_Sname_a1)
    #===============Site Names End===============

    #===============Type Of Sharing===============
    tosA1HostVivaFile = dfVivaHostA1["Type of sharing"].to_list()
    tosA1HostCetinFile = dfCetinHostA1["Type of sharing"].to_list()
    tosA1GuestVivaFile = dfVivaHostViva["Type of sharing"].to_list()
    tosA1GuestCetinFile = dfCetinHostCetin["Type of sharing"].to_list()

    #Dicts Viva/Cetin ID Host/Guest : TOS -> Viva/Cetin ALL ID: Tos
    dictA1IDHost_Tos_viva = dictCreator(vivaHostA1ID, tosA1HostVivaFile)
    dictA1IDHost_Tos_cetin = dictCreator(cetinHostA1ID, tosA1HostCetinFile)
    dictA1IDGuest_Tos_viva = dictCreator(a1GuestIDhostViva, tosA1GuestVivaFile)
    dictA1IDGuest_Tos_cetin = dictCreator(a1GuestIDhostCetin, tosA1GuestCetinFile)
    dictA1ID_TOS_viva = {}
    dictA1ID_TOS_viva.update(dictA1IDHost_Tos_viva)
    dictA1ID_TOS_viva.update(dictA1IDGuest_Tos_viva)
    dictA1ID_TOS_cetin = {}
    dictA1ID_TOS_cetin.update(dictA1IDHost_Tos_cetin)
    dictA1ID_TOS_cetin.update(dictA1IDGuest_Tos_cetin)
    #===============Type Of Sharing End===============

    #===============Collocation Status===============
    statA1HostVivaFile = dfVivaHostA1["Collocation status"].to_list()
    statA1HostCetinFile = dfCetinHostA1["Collocation status"].to_list()
    statA1GuestVivaFile = dfVivaHostViva["Collocation status"].to_list()
    statA1GuestCetinFile = dfCetinHostCetin["Collocation status"].to_list()
    dictA1HostStatVivaFile = dictCreator(vivaHostA1ID, statA1HostVivaFile)
    dictA1GuestStatVivaFile = dictCreator(a1GuestIDhostViva, statA1GuestVivaFile)
    dictA1HostStatCetinFile = dictCreator(cetinHostA1ID, statA1HostCetinFile)
    dictA1GuestStatCetinFile = dictCreator(a1GuestIDhostCetin, statA1GuestCetinFile)

    dictA1StatViva = {}
    dictA1StatViva.update(dictA1HostStatVivaFile)
    dictA1StatViva.update(dictA1GuestStatVivaFile)
    dictA1StatCetin = {}
    dictA1StatCetin.update(dictA1HostStatCetinFile)
    dictA1StatCetin.update(dictA1GuestStatCetinFile)
    #===============Collocation Status End===============

    #===============1'st Guest===============
    guest1CetinHostCetin = dfCetinHostCetin["1's guest"].to_list()
    guest1VivaHostViva = dfVivaHostViva["1's guest"].to_list()
    guest1CetinHostA1 = dfCetinHostA1["1's guest"].to_list()
    guest1VivaHostA1 = dfVivaHostA1["1's guest"].to_list()

    dictA1Host_Guest1_viva = dictCreator(vivaHostA1ID, guest1VivaHostA1)
    dictA1Host_Guest1_cetin = dictCreator(cetinHostA1ID, guest1CetinHostA1)
    dictA1Guest_Guest1_viva = dictCreator(a1GuestIDhostViva, guest1VivaHostViva)
    dictA1Guest_Guest1_cetin = dictCreator(a1GuestIDhostCetin, guest1CetinHostCetin)
    dictList = [dictA1Host_Guest1_viva, dictA1Host_Guest1_cetin, dictA1Guest_Guest1_viva, dictA1Guest_Guest1_cetin]
    for dict in dictList:
        dictUniqueA1ID_Guest1 = dictIntersector(dictUniqueA1ID_Guest1, dict)

    #=====================================================================================

    dfResult = pd.DataFrame (columns=[
            'A1 Site ID',
            'A1 Site Name',
            'Site Owner',
            'Owner',
            'Vivacom Site ID',
            'Vivacom Site Name',
            'Type Of Sharing Vivacom',
            'Collocation Status Vivacom',
            'CETIN Site ID',
            'CETIN Site Name',
            'Type Of Sharing CETIN',
            'Collocation Status CETIN',
            '1\'s Guest',
            'TosFirstGuest;TosSecondGuest'
            ])

# ------------------------------------------------------------------------------ start
    if dismantledFilter == True:
        dfResultDismantled_Vivacom = pd.DataFrame (columns=[
                'A1 Site ID',
                'A1 Site Name',
                'Site Owner',
                'Owner',
                'Vivacom Site ID',
                'Vivacom Site Name',
                'Type Of Sharing Vivacom',
                'Collocation Status Vivacom',
                'CETIN Site ID',
                'CETIN Site Name',
                'Type Of Sharing CETIN',
                'Collocation Status CETIN',
                '1\'s Guest',
                'TosFirstGuest;TosSecondGuest'
                ])
        dfResultDismantled_Cetin = pd.DataFrame (columns=[
                'A1 Site ID',
                'A1 Site Name',
                'Site Owner',
                'Owner',
                'Vivacom Site ID',
                'Vivacom Site Name',
                'Type Of Sharing Vivacom',
                'Collocation Status Vivacom',
                'CETIN Site ID',
                'CETIN Site Name',
                'Type Of Sharing CETIN',
                'Collocation Status CETIN',
                '1\'s Guest',
                'TosFirstGuest;TosSecondGuest'
                ])
# ------------------------------------------------------------------------------ end

    for key, value in dictUniqueA1ID.items():
        statViva = None
        statCetin = None
        cetinSname = None
        vivaSname = None
        tosCetin = None
        tosViva = None
        siteOwner = None
        tosFirstGuest_SecondGuest = None
        tosFirstGuest = None
        tosSecondGuest = None
        a1ID = key
        if len(value)>2:
            cetinID = []
            vivaID = []
        else:
            cetinID = None
            vivaID = None
     
        for item in value:
                if item.isnumeric():
                    if len(value) > 2:
                        cetinID.append(item)
                    else: cetinID = int(item)
                elif re.findall(vivacomRegexPattern,item):
                    if len(value) > 2:
                        vivaID.append(item)
                    else: vivaID = item
                else:
                    for k, v in dictA1GuestIDhostViva_vivaIDhostViva.items():
                        if a1ID == k and v == item:
                            vivaID = item
                    for k, v in dictA1GuestIDhostCetin_cetinHostCetinID.items():
                        if a1ID == k and v == item:
                            cetinID = item
                    for k, v in dictA1GuestIDhostViva_cetinGuestIDhostViva.items():
                        if a1ID == k and v == item:
                            cetinID = item
                    for k, v in dictA1GuestIDhostCetin_vivaGuestIDhostCetin.items():
                        if a1ID == k and v == item:
                            vivaID = item
                    for k, v in dictFileCetinA1ID_guestCetinIDhostA1.items():
                        if a1ID == k and v == item:
                            cetinID = item
                    for k, v in dictFileCetinA1ID_guestVivaIDhostA1.items():
                        if a1ID == k and v == item:
                            vivaID = item
                    for k, v in dictFileVivaA1ID_guestCetinIDhostA1.items():
                        if a1ID == k and v == item:
                            cetinID = item
                    for k, v in dictFileVivaA1ID_guestVivaIDhostA1.items():
                        if a1ID == k and v == item:
                            vivaID = item

        if vivaID != None:
            if isinstance(vivaID, type(list)):
                if len(vivaID) == 1:
                    vivaID = vivaID[0]
        if cetinID != None:
            if isinstance(cetinID, type(list)):
                if len(cetinID) == 1:
                    cetinID = cetinID[0]

        for k, v in dictVivaID_Sname.items():
            if v == "nan":
                v = None
            if isinstance(vivaID, type(list)):
                if vivaID[0] == k:
                    vivaSname = v
                if len(vivaID)>1:
                    if vivaID[1] == k:
                        vivaSname = v
            else:
                if vivaID == k:
                    vivaSname = v

        for k, v in dictCetinID_Sname.items():
            k = int(k)
            if v == "nan":
                v = None
            if isinstance(cetinID, type(list)):
                if int(cetinID[0]) == k:
                    cetinSname = v
                if len(cetinID)>1:
                    if cetinID[1] == k:
                        cetinID = v
            else:
                if cetinID == k:
                    cetinSname = v

        for k, v in dictA1ID_TOS_viva.items():
            if a1ID == k:
                tosViva = v

        for k, v in dictA1ID_TOS_cetin.items():
            if a1ID == k:
                tosCetin = v

        for k, v in dictA1StatViva.items():
            if a1ID == k:
                statViva = v

        for k, v in dictA1StatCetin.items():
            if a1ID == k:
                statCetin = v

        for k, v in dictUniqueA1ID_Owner.items():
            if a1ID == k:
                a1Owner = v

        for k, v in dictUniqueA1ID_Sname.items():
            if a1ID == k:
                a1Sname = v
     
        for k, v in dictUniqueA1ID_Guest1.items():
            if a1ID == k:
                if v.isnumeric():
                    guest = int(v)
                else: guest = v

        for k,v in dictUniqueA1ID_SiteOwner.items():
            if a1ID == k:
                siteOwner = v
        if isinstance(cetinID, type(list)) and len(cetinID) == 0:
            cetinID = None
        if isinstance(vivaID, type(list)) and len(vivaID) == 0:
            vivaID = None
# ------------------------------------------------------------------------------ start
        if type(guest) == str:
            if re.findall(a1RegexPattern, guest):
                if siteOwner == "CETIN":
                    if tosCetin is not None:
                        tosFirstGuest = tosCetin
                    else: tosFirstGuest = "N/A"
                    if tosViva is not None:
                        tosSecondGuest = tosViva
                elif siteOwner == "Vivacom":
                    if tosViva is not None:
                        tosFirstGuest = tosViva
                    else: tosFirstGuest = "N/A"
                    if tosCetin is not None:
                        tosSecondGuest = tosCetin
            else:
                if tosViva is not None:
                    tosFirstGuest = tosViva
                else: tosFirstGuest = "N/A"
                if tosCetin is not None:
                    tosSecondGuest = tosCetin
        else:
            if tosCetin is not None:
                tosFirstGuest = tosCetin
            else: tosFirstGuest = "N/A"
            if tosViva is not None:
                tosSecondGuest = tosViva
        if tosSecondGuest is not None:
            tosFirstGuest_SecondGuest = tosFirstGuest + " ; " + tosSecondGuest
        else: tosFirstGuest_SecondGuest = tosFirstGuest + "; " + "N/A"

        if dismantledFilter == True:
            if statViva == "Dismantled" or statViva == "Stopped":
                '''
                dfDismantled_Stopped_Vivacom = pd.DataFrame({'A1 Site ID': [a1ID],
                                'A1 Site Name': [a1Sname],
                                'Owner': [a1Owner],
                                'Vivacom Site ID': [vivaID],
                                'Vivacom Site Name': [vivaSname],
                                'Type Of Sharing Vivacom': [tosViva],
                                'Collocation Status Vivacom': [statViva],
                                'CETIN Site ID': [cetinID],
                                'CETIN Site Name': [cetinSname],
                                'Type Of Sharing CETIN': [tosCetin],
                                'Collocation Status CETIN':[statCetin],
                                '1\'s Guest': [guest],
                                'Site Owner': [siteOwner],
                                'TosFirstGuest;TosSecondGuest': [tosFirstGuest_SecondGuest]})
                dfResultDismantled_Vivacom = pd.concat([dfResultDismantled_Vivacom, dfDismantled_Stopped_Vivacom], ignore_index=True, axis=0)
                '''
                statViva = None
                vivaID = None
                tosViva = None
                vivaSname = None
            if statCetin == "Dismantled" or statCetin == "Stopped":
                '''
                dfDismantled_Stopped_Cetin = pd.DataFrame({'A1 Site ID': [a1ID],
                                'A1 Site Name': [a1Sname],
                                'Owner': [a1Owner],
                                'Vivacom Site ID': [vivaID],
                                'Vivacom Site Name': [vivaSname],
                                'Type Of Sharing Vivacom': [tosViva],
                                'Collocation Status Vivacom': [statViva],
                                'CETIN Site ID': [cetinID],
                                'CETIN Site Name': [cetinSname],
                                'Type Of Sharing CETIN': [tosCetin],
                                'Collocation Status CETIN':[statCetin],
                                '1\'s Guest': [guest],
                                'Site Owner': [siteOwner],
                                'TosFirstGuest;TosSecondGuest': [tosFirstGuest_SecondGuest]})
                dfResultDismantled_Cetin = pd.concat([dfResultDismantled_Cetin, dfDismantled_Stopped_Cetin], ignore_index=True, axis=0)
                '''
                statCetin = None
                cetinID = None
                tosCetin = None
                cetinSname = None
# ------------------------------------------------------------------------------ end

        singleRow = pd.DataFrame({
                            'A1 Site ID': [a1ID],
                            'A1 Site Name': [a1Sname],
                            'Owner': [a1Owner],
                            'Vivacom Site ID': [vivaID],
                            'Vivacom Site Name': [vivaSname],
                            'Type Of Sharing Vivacom': [tosViva],
                            'Collocation Status Vivacom': [statViva],
                            'CETIN Site ID': [cetinID],
                            'CETIN Site Name': [cetinSname],
                            'Type Of Sharing CETIN': [tosCetin],
                            'Collocation Status CETIN':[statCetin],
                            '1\'s Guest': [guest],
                            'Site Owner': [siteOwner],
                            'TosFirstGuest;TosSecondGuest': [tosFirstGuest_SecondGuest]})
# ------------------------------------------------------------------------------ start
        if dismantledFilter is True:
            if (vivaID is None) and (cetinID is None): #drop the whole row if both id's are None
                singleRow.iloc[0:0]
# ------------------------------------------------------------------------------ end

        dfResult = pd.concat([dfResult, singleRow], ignore_index = True, axis = 0)
    #dfResult['A1 Site ID'] = dfResult['A1 Site ID'].str.strip()#премахнат
    writer = pd.ExcelWriter('C:\\Users\\a1bg511027\\Desktop\\ResultPy.xlsx',engine = 'xlsxwriter') #добавен engine

    dfResult.sort_values('A1 Site ID', inplace = True)
    dfResult.reset_index(drop = True, inplace=True)
    dfResult.set_index('A1 Site ID',inplace = True)
    dfStylerResult = dfResult.style.set_properties(**{'text-align': 'left'})
    dfStylerResult.to_excel(writer, sheet_name='On Air')

# ------------------------------------------------------------------------------ start
    '''
    if dismantledFilter == True:
        dfResultDismantled_Vivacom.sort_values('A1 Site ID', inplace = True)
        dfResultDismantled_Vivacom.reset_index(drop = True, inplace=True)
        dfResultDismantled_Vivacom.set_index('A1 Site ID',inplace = True)
        dfStylerDismantled_Vivacom = dfResultDismantled_Vivacom.style.set_properties(**{'text-align': 'left'})
        dfStylerDismantled_Vivacom.to_excel(writer, sheet_name='Vivacom Dismantled')


        dfResultDismantled_Cetin.sort_values('A1 Site ID', inplace = True)
        dfResultDismantled_Cetin.reset_index(drop = True, inplace=True)
        dfResultDismantled_Cetin.set_index('A1 Site ID',inplace = True)
        dfStylerDismantled_Cetin = dfResultDismantled_Cetin.style.set_properties(**{'text-align': 'left'})
        dfStylerDismantled_Cetin.to_excel(writer, sheet_name='Cetin Dismantled')
    '''
# ------------------------------------------------------------------------------ end

    writer.save()
    '''
    'A1 Site ID' #=====check=====
    'A1 Site name' #=====check=====
    'Owner' #=====check=====
    'Guest 1' #=====check=====
    'Vivacom Site ID' #=====check=====
    'Vivacom Site name' #=====check=====
    'Type of Sharing Vivacom' #=====check=====
    'Collocation status Vivacom' #=====check=====
    'Activated on Vivacom' ??   
    'CETIN Site ID' #=====check=====
    'CETIN Site name' #=====check=====
    'Type of Sharing CETIN' #=====check=====
    'Collocation status CETIN' #=====check=====
    'Activated on CETIN' ??
    '''