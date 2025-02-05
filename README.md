# autolink
To use this script, make sure you curate a list of game titles as close to the original title as possible in a separate txt file called "input.txt" and place it in a folder titled input located at the same level as the script.

The program will prompt you for a url, a destination where you want the files to be downloaded to, and if you want to cross reference a local folder location for list matches to avoid checking the url for those titles. It will then scrape the url content for link matches, clean the link titles and the input item titles and compare them for matches. If a match exists, it will then check if the file already exists and if not, download it.

You will have the opportunity to crawl the list over and over across other endpoints as you narrow it down per url without having to enter file related information again.
