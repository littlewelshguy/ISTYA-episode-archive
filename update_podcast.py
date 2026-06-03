import pandas as pd
import requests
import xml.etree.ElementTree as ET
import re
import os

# CONFIGURATION
RSS_URL = "https://istya.libsyn.com/rss"
EXCEL_FILE = "I_Saw_That_Years_Ago_Comprehensive_Archive.xlsx"
OMDB_API_KEY = "YOUR_API_KEY_HERE" # Replace with your key from http://www.omdbapi.com/

def get_latest_episodes():
    print("Fetching latest episodes from RSS...")
    try:
        # Use a browser-like User-Agent to avoid blocks
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(RSS_URL, headers=headers)
        
        # Simple regex extraction as backup for XML parsing
        content = response.text
        items = re.findall(r'<item>(.*?)</item>', content, re.DOTALL)
        
        episodes = []
        for item in items:
            title_match = re.search(r'<title>(.*?)</title>', item)
            pub_date_match = re.search(r'<pubDate>(.*?)</pubDate>', item)
            link_match = re.search(r'<link>(.*?)</link>', item)
            
            title = title_match.group(1) if title_match else ""
            pub_date = pub_date_match.group(1) if pub_date_match else ""
            link = link_match.group(1) if link_match else ""
            
            # Clean CDATA
            title = re.sub(r'<!\[CDATA\[(.*?)\]\]>', r'\1', title)
            
            match = re.search(r'Ep\s+\d+\s+-\s+(.*?)\s*\((\d{4})\)', title)
            film_name = match.group(1).strip() if match else title
            film_year = match.group(2).strip() if match else ""
            
            episodes.append({
                'Episode Title': title,
                'Publication Date': pub_date,
                'Film Name': film_name,
                'Film Year': film_year,
                'Podcast Link': link
            })
        return pd.DataFrame(episodes)
    except Exception as e:
        print(f"Error fetching RSS: {e}")
        return None

def enrich_with_omdb(df):
    if OMDB_API_KEY == "YOUR_API_KEY_HERE":
        print("Skipping OMDb enrichment (no API key provided).")
        return df
    
    print("Enriching with OMDb data...")
    new_cols = [
        'Genre', 'IMDb Rating', 'IMDb Votes', 'Rotten Tomatoes', 'Metascore', 
        'Director', 'Writer', 'Actors', 'Plot', 'Language', 'Country', 
        'Awards', 'Runtime', 'Released', 'Rated', 'Type', 'BoxOffice', 'Production'
    ]
    
    for col in new_cols:
        if col not in df.columns:
            df[col] = "Pending"
            
    for index, row in df.iterrows():
        # Only fetch if data is missing or marked as "Pending/Use Update Tool"
        if pd.isna(row.get('Genre')) or row.get('Genre') in ["Use Update Tool", "Pending"]:
            title = row['Film Name']
            year = row['Film Year']
            print(f"Fetching data for: {title} ({year})")
            url = f"http://www.omdbapi.com/?t={title}&y={year}&plot=full&apikey={OMDB_API_KEY}"
            try:
                res = requests.get(url).json()
                if res.get('Response') == 'True':
                    df.at[index, 'Genre'] = res.get('Genre')
                    df.at[index, 'IMDb Rating'] = res.get('imdbRating')
                    df.at[index, 'IMDb Votes'] = res.get('imdbVotes')
                    df.at[index, 'Metascore'] = res.get('Metascore')
                    df.at[index, 'Director'] = res.get('Director')
                    df.at[index, 'Writer'] = res.get('Writer')
                    df.at[index, 'Actors'] = res.get('Actors')
                    df.at[index, 'Plot'] = res.get('Plot')
                    df.at[index, 'Language'] = res.get('Language')
                    df.at[index, 'Country'] = res.get('Country')
                    df.at[index, 'Awards'] = res.get('Awards')
                    df.at[index, 'Runtime'] = res.get('Runtime')
                    df.at[index, 'Released'] = res.get('Released')
                    df.at[index, 'Rated'] = res.get('Rated')
                    df.at[index, 'Type'] = res.get('Type')
                    df.at[index, 'BoxOffice'] = res.get('BoxOffice')
                    df.at[index, 'Production'] = res.get('Production')
                    
                    rt = "N/A"
                    for r in res.get('Ratings', []):
                        if r['Source'] == 'Rotten Tomatoes': rt = r['Value']
                    df.at[index, 'Rotten Tomatoes'] = rt
            except Exception as e:
                print(f"Error for {title}: {e}")
            time.sleep(0.1) # Respect API limits
            
    return df

def update():
    new_df = get_latest_episodes()
    if new_df is None: return
    
    if os.path.exists(EXCEL_FILE):
        old_df = pd.read_excel(EXCEL_FILE, sheet_name='Episodes')
        # Ensure all columns exist in old_df
        combined = pd.concat([new_df, old_df]).drop_duplicates(subset=['Episode Title'], keep='last')
    else:
        combined = new_df
        
    combined = enrich_with_omdb(combined)
    
    with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
        combined.to_excel(writer, sheet_name='Episodes', index=False)
        # Add a basic guide sheet
        pd.DataFrame({'Info': ['Last updated: ' + pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')]}).to_excel(writer, sheet_name='Log', index=False)
    
    print(f"Update complete. Total episodes in archive: {len(combined)}")

if __name__ == "__main__":
    import time
    update()
