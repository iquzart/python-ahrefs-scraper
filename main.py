import yaml
import pandas as pd
import download_csv
import scraper
import re


def main():

    # Load configuration file
    with open("config.yaml", "r") as ymlfile:
        cfg = yaml.safe_load(ymlfile)

    # Login th AHREFS URL
    download_csv.login(cfg)

    url_info = []
    j = 0
    for group in cfg: 
        if (j >= 3): 
            urls = cfg[group]
            for target_url in urls:
                if re.match(r'^#', target_url):
                    split = target_url.split("#")
                    url = ' '.join(split)
                    #print ("formated URL: ",url)
                    # Download CSV files
                    dr, ur = download_csv.export_csv(url)
                else:
                    # Download CSV files
                    print (target_url)
                    dr, ur = download_csv.export_csv(target_url)
                data = {group: {'url': target_url,'ur': ur,'dr': dr}}
                url_info.append(data)
                
        j = j + 1
    # Logout from AHREFS
    download_csv.logout(cfg)
       
    # Create DataFrame from the Downliaded CSVs
    scrap_df = scraper.exstract_data(url_info)
    
    # Split DF and export to Excel based on URL group
    scraper.export_to_excel(scrap_df, cfg)
  

if __name__== "__main__":
  main()