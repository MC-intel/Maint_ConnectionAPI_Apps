# Run function and find which units in the Procal report are in Maintenance Connection by merging the two dataframes and making a list of matching values.
GETdf = pullFromAPI()

merged_df = pd.merge(GETdf, result_df[['M', 'nextcal', 'caldates']], how='inner', on='M')
final_df = merged_df[['PK', 'Name', 'M', 'nextcal', 'caldates']]

print('\nUpdating the following Equipment: ')
print(final_df)
final_df.to_csv(os.path.join(desktop_dir2, 'updatedUnits.csv'), index=False)
time.sleep(5)
print(' ')
