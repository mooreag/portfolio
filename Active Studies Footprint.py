#!/usr/bin/env python
# coding: utf-8

# This script takes active studies, separates export into tabs for each executing party, then pivots those tabs to give output. A sum of the total number of patients enrolled across all studies for each site is also added. **NOTE:** If you change any column titles in the query, you will have to update them here too.
# 
# 1. Run **STEP 1** section
# 2. Depending on which executing parties you filtered for in your query or if you want to even split them into separate tabs, you will either use the **Unfiltered** section (which will not split the studies out) or the **Filter: executing_party** section(s) (which will split the studies out into a tab for each executing party)
# 3. Once you run the proper sections, read those tab/s into an excel doc

# # 1. STEP 1

# In[ ]:


import pandas as pd


# In[ ]:


df = pd.read_excel('active_studies_export.xlsx')


# # Unfiltered
# This will pivot the entire query export without splitting it out into executing_party tabs. If you don't need the different BUs to be filtered, then run this one.

# In[ ]:


# group total enrolled
ser1 = df.groupby(['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site']).total_enrolled.sum().to_frame()
ser1


# In[ ]:


# group sad date
ser2 = df.groupby(['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site']).sad.min().to_frame()


# In[ ]:


# group site status
ser3 = df.groupby(['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site']).site_status.min().to_frame()


# In[ ]:


# remerge the 3 groups
merge1 = pd.merge(ser1, ser2,  how='outer', left_on=['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'], right_on = ['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'])


# In[ ]:


merge2 = pd.merge(merge1, ser3,  how='outer', left_on=['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'], right_on = ['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'])
merge2


# In[ ]:


# reset index
modified = merge2.reset_index()


# In[ ]:


# convert sad datetime to string to remove the time from the final output
modified['sad'] = modified['sad'].astype(str)
# modified


# In[ ]:


# pivot transformation (https://stackoverflow.com/questions/55416191/pandas-change-order-of-columns-in-pivot-table)
pivot = (modified.pivot_table(index=['site'],
                             columns=['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall'],
                             values=['site_status', 'sad','total_enrolled'],
                             aggfunc=lambda x: ' '.join(str(v) for v in x)).sort_index(axis=1, level=1))


# In[ ]:


# convert total_enrolled to float (https://stackoverflow.com/questions/48094854/python-convert-object-to-float)
pivot['total_enrolled'] = pivot['total_enrolled'].astype(float)
# pivot['total_enrolled']
pivot


# In[ ]:


# reorder column titles (https://github.com/pandas-dev/pandas/issues/4720)
pivot.columns = pivot.columns.swaplevel(7, 0).swaplevel(6, 0).swaplevel(5, 0).swaplevel(4, 0).swaplevel(3, 0).swaplevel(2, 0).swaplevel(1, 0)
pivot.sort_index(1)


# In[ ]:


# sum enrollment across columns
pivot['total_patients_recruited'] = pivot.sum(axis=1)
# pivot


# In[ ]:


# print to excel
pivot.to_excel('pivot_export.xlsx')


# # Filter: executingparty1
# Creates a tab just for executingparty1 studies

# In[ ]:


# filter for executingparty1
executingparty1 = df.loc[df['executing_party'] == 'executingparty1']
executingparty1 = executingparty1.reset_index(drop=True)


# In[ ]:


# group total enrolled
ser1_executingparty1 = executingparty1.groupby(['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site']).total_enrolled.sum().to_frame()


# In[ ]:


# group sad date
ser2_executingparty1 = executingparty1.groupby(['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site']).sad.min().to_frame()


# In[ ]:


# group site status
ser3_executingparty1 = executingparty1.groupby(['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site']).site_status.min().to_frame()


# In[ ]:


# remerge the 3 groups
merge1_executingparty1 = pd.merge(ser1_executingparty1, ser2_executingparty1,  how='outer', left_on=['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'], right_on = ['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'])


# In[ ]:


merge2_executingparty1 = pd.merge(merge1_executingparty1, ser3_executingparty1,  how='outer', left_on=['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'], right_on = ['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'])


# In[ ]:


# reset index
modified_executingparty1 = merge2_executingparty1.reset_index()


# In[ ]:


# convert sad datetime to string to remove the time from the final output
modified_executingparty1['sad'] = modified_executingparty1['sad'].astype(str)
# modified_executingparty1


# In[ ]:


# pivot transformation (https://stackoverflow.com/questions/55416191/pandas-change-order-of-columns-in-pivot-table)
pivot_executingparty1 = (modified_executingparty1.pivot_table(index=['site'],
                             columns=['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall'],
                             values=['site_status', 'sad','total_enrolled'],
                             aggfunc=lambda x: ' '.join(str(v) for v in x)).sort_index(axis=1, level=1))


# In[ ]:


# convert total_enrolled to float (https://stackoverflow.com/questions/48094854/python-convert-object-to-float)
pivot_executingparty1['total_enrolled'] = pivot_executingparty1['total_enrolled'].astype(float)
# pivot['total_enrolled']
# pivot_executingparty1


# In[ ]:


# reorder column titles (https://github.com/pandas-dev/pandas/issues/4720)
pivot_executingparty1.columns = pivot_executingparty1.columns.swaplevel(7, 0).swaplevel(6, 0).swaplevel(5, 0).swaplevel(4, 0).swaplevel(3, 0).swaplevel(2, 0).swaplevel(1, 0)
pivot_executingparty1.sort_index(1)


# In[ ]:


# sum enrollment across columns
pivot_executingparty1['total_patients_recruited'] = pivot_executingparty1.sum(axis=1)
# pivot_executingparty1


# In[ ]:


# pivot_executingparty1.to_excel('pivot_export_executingparty1.xlsx')


# # Filter: executingparty2
# Creates a tab just for executingparty2 studies

# In[ ]:


# filter for executingparty2
executingparty2 = df.loc[df['executing_party'] == 'executingparty2']
executingparty2 = executingparty2.reset_index(drop=True)


# In[ ]:


# group total enrolled
ser1_executingparty2 = executingparty2.groupby(['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site']).total_enrolled.sum().to_frame()


# In[ ]:


# group sad date
ser2_executingparty2 = executingparty2.groupby(['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site']).sad.min().to_frame()


# In[ ]:


# group site status
ser3_executingparty2 = executingparty2.groupby(['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site']).site_status.min().to_frame()


# In[ ]:


# remerge the 3 groups
merge1_executingparty2 = pd.merge(ser1_executingparty2, ser2_executingparty2,  how='outer', left_on=['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'], right_on = ['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'])


# In[ ]:


merge2_executingparty2 = pd.merge(merge1_executingparty2, ser3_executingparty2,  how='outer', left_on=['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'], right_on = ['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'])


# In[ ]:


# reset index
modified_executingparty2 = merge2_executingparty2.reset_index()


# In[ ]:


# convert sad datetime to string to remove the time from the final output
modified_executingparty2['sad'] = modified_executingparty2['sad'].astype(str)
# modified_executingparty2


# In[ ]:


# pivot transformation (https://stackoverflow.com/questions/55416191/pandas-change-order-of-columns-in-pivot-table)
pivot_executingparty2 = (modified_executingparty2.pivot_table(index=['site'],
                             columns=['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall'],
                             values=['site_status', 'sad','total_enrolled'],
                             aggfunc=lambda x: ' '.join(str(v) for v in x)).sort_index(axis=1, level=1))


# In[ ]:


# convert total_enrolled to float (https://stackoverflow.com/questions/48094854/python-convert-object-to-float)
pivot_executingparty2['total_enrolled'] = pivot_executingparty2['total_enrolled'].astype(float)
# pivot['total_enrolled']
# pivot_executingparty2


# In[ ]:


# reorder column titles (https://github.com/pandas-dev/pandas/issues/4720)
pivot_executingparty2.columns = pivot_executingparty2.columns.swaplevel(7, 0).swaplevel(6, 0).swaplevel(5, 0).swaplevel(4, 0).swaplevel(3, 0).swaplevel(2, 0).swaplevel(1, 0)
pivot_executingparty2.sort_index(1)


# In[ ]:


# sum enrollment across columns
pivot_executingparty2['total_patients_recruited'] = pivot_executingparty2.sum(axis=1)
# pivot_executingparty2


# # Filter: executingparty2
# Creates a tab just for executingparty2 studies

# In[ ]:


# filter for executingparty2
executingparty2 = df.loc[df['executing_party'] == 'executingparty2']
executingparty2 = executingparty2.reset_index(drop=True)


# In[ ]:


# group total enrolled
ser1_executingparty2 = executingparty2.groupby(['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site']).total_enrolled.sum().to_frame()


# In[ ]:


# group sad date
ser2_executingparty2 = executingparty2.groupby(['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site']).sad.min().to_frame()


# In[ ]:


# group site status
ser3_executingparty2 = executingparty2.groupby(['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site']).site_status.min().to_frame()


# In[ ]:


# remerge the 3 groups
merge1_executingparty2 = pd.merge(ser1_executingparty2, ser2_executingparty2,  how='outer', left_on=['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'], right_on = ['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'])


# In[ ]:


merge2_executingparty2 = pd.merge(merge1_executingparty2, ser3_executingparty2,  how='outer', left_on=['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'], right_on = ['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall','site'])


# In[ ]:


# reset index
modified_executingparty2 = merge2_executingparty2.reset_index()


# In[ ]:


# convert sad datetime to string to remove the time from the final output
modified_executingparty2['sad'] = modified_executingparty2['sad'].astype(str)
# modified_executingparty2


# In[ ]:


# pivot transformation (https://stackoverflow.com/questions/55416191/pandas-change-order-of-columns-in-pivot-table)
pivot_executingparty2 = (modified_executingparty2.pivot_table(index=['site'],
                             columns=['study_number','executing_party','phase','molecule','fpa_year','study_fpi_year','firewall'],
                             values=['site_status', 'sad','total_enrolled'],
                             aggfunc=lambda x: ' '.join(str(v) for v in x)).sort_index(axis=1, level=1))


# In[ ]:


# convert total_enrolled to float (https://stackoverflow.com/questions/48094854/python-convert-object-to-float)
pivot_executingparty2['total_enrolled'] = pivot_executingparty2['total_enrolled'].astype(float)
# pivot['total_enrolled']
# pivot_executingparty2


# In[ ]:


# reorder column titles (https://github.com/pandas-dev/pandas/issues/4720)
pivot_executingparty2.columns = pivot_executingparty2.columns.swaplevel(7, 0).swaplevel(6, 0).swaplevel(5, 0).swaplevel(4, 0).swaplevel(3, 0).swaplevel(2, 0).swaplevel(1, 0)
pivot_executingparty2.sort_index(1)


# In[ ]:


# sum enrollment across columns
pivot_executingparty2['total_patients_recruited'] = pivot_executingparty2.sum(axis=1)
# pivot_executingparty2


# # save all 3 tabs to the same excel doc

# In[ ]:


writer = pd.ExcelWriter('Active Studies Footprint_[date].xlsx')

pivot_executingparty1.to_excel(writer, sheet_name = 'executingparty1')
pivot_executingparty2.to_excel(writer, sheet_name = 'executingparty2')
pivot_executingparty2.to_excel(writer, sheet_name = 'executingparty2')

writer.save()

