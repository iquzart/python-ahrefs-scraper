import pandas as pd
import numpy as np
import re
import string
import os
import glob



def exstract_data(url_info):
    '''
    Create DataFrame from the Downliaded CSVs
    '''
    print ("Creating DataFrame")
    output = pd.DataFrame(columns=[
                    'group', 'URLs', 'DR','UR','DR Buckets:','10+', '30+', '45+', '60+', 'Tier Buckets:', '0', '1-3', '4-10', '11+'])
    
    # Chrome Download path
    filepath = 'C:\\Users\\iqbal\\Downloads'
    print ("CSV Donwload Location", filepath)

    j = 1
    for d in url_info:
        for group, properties in d.items():
            url = properties['url']
            dr = properties['dr']
            ur = properties['ur']
            cut_string = url.split('/')   
            domain_name = cut_string[2]

            os.chdir( filepath )
            files = glob.glob( domain_name + '*.csv' )

            
            for filename in files:
                input = pd.read_csv(filename)
                print ("CSV loaded: ", filename)
                data = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
                data[0] = group
                data[1] = url
                input = input[input['Type'].str.contains("Dofollow")]
                input = input[input['Linked Domains'] <= 100]

                data[2] = dr
                data[3] = ur

                for row in input.itertuples():
                    domain_rating = row[3]
                    if domain_rating is None:
                        break
                    if 10 <= domain_rating <= 29:
                        data[5] += 1
                    elif 30 <= domain_rating <= 44:
                        data[6] += 1
                    elif 45 <= domain_rating <= 59:
                        data[7] += 1
                    elif 60 <= domain_rating <= 100:
                        data[8] += 1     
        
                    referring_domains = row[5]
                    if referring_domains is None:
                        break
                    if 0 == referring_domains:
                        data[10] += 1
                    if 1 <= referring_domains <= 3:
                        data[11] += 1
                    if 4 <= referring_domains <= 10:
                        data[12] += 1
                    if 11 <= referring_domains:
                        data[13] += 1     
                j += 1
                output.loc[len(output)] = data
                output.dropna(inplace=True)
                output = output.append(
                    pd.Series(dtype='float64'), ignore_index=True)
    print ("DataFrame is Ready!")

    # return DF
    return output

        
def export_to_excel(df, cfg):
    '''
    Split DF and export to Excel based on URL group
    '''
    k = 0
    for url_group in cfg: 
        if (k >= 3): 
          
            # Format Group Name
            split = url_group.split("_")
            join = ' '.join(split)
            formatted_group_name =  (string.capwords(join))
            
            print ('Exporting DataFrame to Excel for URL Group: ' + formatted_group_name)

            # grouped DF
            df_groupd = df.groupby(df.group)
            
            # Get grouped data from DF
            group = df_groupd.get_group(url_group)

            # Copy DF
            group = group.copy()

            # Replace URLs header with URL Group name
            group.rename(columns={'URLs': formatted_group_name}, inplace=True)

            # Remove Group column
            group.drop('group', axis=1, inplace=True)
            
            # set index
            group.insert(0, '', range(1, 1 + len(group)))

            # set index prefix
            group[''] = '#' + group[''].astype(str)

            # set index suffix
            group[''] = group[''].astype(str) + ' Competitor:'

            # Empty Buckets
            group['DR Buckets:']=''
            group['Tier Buckets:']=''
                    
            #Insert Empty Row
            group_nan = pd.DataFrame([[np.nan] * len(group.columns)], columns=group.columns)
            group = group_nan.append(group, ignore_index=True)

            
            # Money Site checker
            if re.match(r'^#', group.iloc[1,1]): 
    
                # Shift Columns stating from URL
                mask = ~(group.columns.isin(['']))
                cols_to_shift = group.columns[mask]
                group[cols_to_shift] = group[cols_to_shift].shift(-1)
                #Add MS Name to id column
                group.iloc[0,0]="Money Site: "

                # Remove '#' from MS URL
                group = group.replace('#http', 'http', regex=True)

                # Drop NaN at end
                group.dropna(inplace=True)
            else: 
                # Row insert for no Money Site
                group.loc[0] = ['Money Site: ','',0,0,'',0,0,0,0,'',0,0,0,0]
                
            # last colomn
            lc = len(group)+1
            
            # Add Competitor Average
            group_ca = ['','Competitor Average', '=IFERROR(ROUND(AVERAGE(C3:C' + str(lc) +'),1),"")', 
                        '=IFERROR(ROUND(AVERAGE(D3:D' + str(lc) +'),1),"")','','=IFERROR(ROUND(AVERAGE(F3:F' + str(lc) +'),1),"")',
                        '=IFERROR(ROUND(AVERAGE(G3:G' + str(lc) +'),1),"")','=IFERROR(ROUND(AVERAGE(H3:H' + str(lc) +'),1),"")',
                        '=IFERROR(ROUND(AVERAGE(I3:I' + str(lc) +'),1),"")','','=IFERROR(ROUND(AVERAGE(K3:K' + str(lc) +'),1),"")',
                        '=IFERROR(ROUND(AVERAGE(L3:L' + str(lc) +'),1),"")','=IFERROR(ROUND(AVERAGE(M3:M' + str(lc) +'),1),"")',
                        '=IFERROR(ROUND(AVERAGE(N3:N' + str(lc) +'),1),"")']
            # Add Competitor Average - Money Site (Difference)
            group_ca_ms = ['','Competitor Average - Money Site (Difference)','=IFERROR(minus(C' + str(lc+1) +',C2),"")',
                           '=IFERROR(minus(D' + str(lc+1) +',D2),"")','','=IFERROR(minus(F' + str(lc+1) +',F2),"")','=IFERROR(minus(G' + str(lc+1) +',G2),"")',
                           '=IFERROR(minus(H' + str(lc+1) +',H2),"")','=IFERROR(minus(I' + str(lc+1) +',I2),"")','','=IFERROR(minus(K' + str(lc+1) +',K2),"")',
                           '=IFERROR(minus(L' + str(lc+1) +',L2),"")','=IFERROR(minus(M' + str(lc+1) +',M2),"")','=IFERROR(minus(N' + str(lc+1) +',N2),"")']
            
            group.loc[len(group)] = group_ca
            group.loc[len(group)] = group_ca_ms
            
            print (group)

            # Create a Pandas Excel writer using XlsxWriter as the engine.
            writer = pd.ExcelWriter(url_group + '.xlsx', engine='xlsxwriter')

            # Convert the dataframe to an XlsxWriter Excel object.
            group.to_excel(writer, sheet_name='Sheet1', index=False)

            # Get workbook
            workbook = writer.book

            # Get Sheet1
            worksheet = writer.sheets['Sheet1']

            # add format for headers
            header_format = workbook.add_format()
            header_format.set_font_name('Arial')
            header_format.set_font_color('white')
            header_format.set_font_size(10)
            header_format.set_fg_color('368de3')
            header_format.set_align('center')

            
            # Set the column width and format.
            worksheet.set_column('A:A', 15)       
            worksheet.set_column('B:B', 45)
            worksheet.set_column('C:C', 5,)
            worksheet.set_column('D:D', 5)
            worksheet.set_column('E:E', 18)
            worksheet.set_column('F:F', 5)
            worksheet.set_column('G:G', 5)
            worksheet.set_column('H:H', 5)
            worksheet.set_column('I:I', 5)
            worksheet.set_column('J:J', 18)
            worksheet.set_column('K:K', 5)
            worksheet.set_column('L:L', 5)
            worksheet.set_column('M:M', 5)
            worksheet.set_column('N:N', 5)



            # Write the column headers with the defined format.
            for col_num, value in enumerate(group.columns.values):
                worksheet.write(0, col_num, value, header_format)

            writer.save()
            print ("Successfully exported report to {}.xlsx".format(url_group))
        k = k + 1




    