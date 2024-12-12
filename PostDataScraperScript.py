#%%
import pandas as pd
import praw
import time
import random
from datetime import datetime, timedelta

# Client ID: "nbpJrnxe2bRtRWaetbSKFA"
# Client Secret: "TmJqeuiqTleE2WD2tLCinaon7AJVVA"
# User Agent: "PostDataScraper by Royal-Ad-9570"
reddit = praw.Reddit(
    client_id="nbpJrnxe2bRtRWaetbSKFA", client_secret="TmJqeuiqTleE2WD2tLCinaon7AJVVA", user_agent="PostDataScraper/1.0 by Royal-Ad-9570"
)

one_week_ago = datetime.utcnow() - timedelta(weeks=1)
subreddit = reddit.subreddit("all")

# ONLY UNCOMMENT IF YOU NEED TO RESTART THE PROGRAM WITHOUT REMEMBERING VARIABLES
datapoints = 4773  # CHANGE THIS NUMBER IF RESTARTING SCRIPT AFTER CLOSING
oldval = datapoints + 1
old = oldval - 1
new = 0
error = 0


df = pd.DataFrame(columns=[                    
                "Title",
                "authorVal",
                "authorMod",
                "authorKarma",
                "Num_Comments",
                "subredditVal",
                "Subreddit_Subcribers",
                "created_utcVal",
                "Score",
                "Ratio",
                "NSFW",
                "Spoiler",
                "Awards",
                "MediaEmbed",
                "SecureMediaEmbed",
                "Media",
                "SecureMedia",
                "URL",
                "Text"])

# ONLY UNCOMMENT IF CREATING NEW DATASET
#df.to_excel('data.xlsx', index=False)  # place headers


for i in range(2000):
    df = pd.DataFrame(columns=[                    
                    "Title",
                    "authorVal",
                    "authorMod",
                    "authorKarma",
                    "Num_Comments",
                    "subredditVal",
                    "Subreddit_Subcribers",
                    "created_utcVal",
                    "Score",
                    "Ratio",
                    "NSFW",
                    "Spoiler",
                    "Awards",
                    "MediaEmbed",
                    "SecureMediaEmbed",
                    "Media",
                    "SecureMedia",
                    "URL",
                    "Text"])

    for randPost in [submission for submission in subreddit.top(limit=10)]:
        randPost.comments.replace_more(limit=0)  # remove more comments
        comment = random.choice(randPost.comments.list())  # get a random comment on the post
        author = comment.author  # get the author
        if author:
            try:
                post = random.choice(list((author.submissions.new(limit=3))))  # get their latest post

                if post == None:
                    print("None")
                    error += 1

                post_time = datetime.utcfromtimestamp(post.created_utc)
                if post_time <= one_week_ago:

                    postScore = post.score
                    postRatio = post.upvote_ratio

                    post_data = {
                        "Title": post.title,
                        "authorVal": post.author,
                        "authorMod": post.author.is_mod,
                        "authorKarma": post.author.link_karma,
                        "Num_Comments": post.num_comments,
                        "subredditVal": str(post.subreddit),
                        "Subreddit_Subcribers": post.subreddit_subscribers,
                        "created_utcVal": post.created_utc,
                        "Score": postScore,
                        "Ratio": postRatio,
                        "NSFW": post.over_18,
                        "Spoiler": post.spoiler,
                        "Awards": post.all_awardings,
                        "MediaEmbed": post.media_embed,
                        "SecureMediaEmbed": post.secure_media_embed,
                        "Media": post.media,
                        "SecureMedia": post.secure_media,
                        "URL": post.url,
                        "Text": post.selftext,
                    }
                    
                    df.loc[old] = post_data

                    old += 1
                    #print(post_time)
                elif post_time > one_week_ago:
                    new += 1
                    #print(post_time)
            except Exception as ExceptionError:
                print("An exception occurred:", type(ExceptionError).__name__) # An exception occurred: 
                error += 1
        else:
            error += 1
                



        print("Old:", old)
        print("New:", new)
        print("Error:", error)


    with pd.ExcelWriter('data.xlsx',
                        mode='a', if_sheet_exists='overlay') as writer:  
        df.to_excel(writer, sheet_name='Sheet1', startrow=oldval, header=False, index=False)
    oldval = old + 1

    print("DONE\n")




