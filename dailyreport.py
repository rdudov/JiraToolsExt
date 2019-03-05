import datetime as dt
import pytz


def comment_clause(created, body, shift=1, keyword='#dailyreport'):
    created_date = get_created_date(created)
    now = dt.datetime.now(created_date.tzinfo)
    return created_date + dt.timedelta(days=shift) > now and body.lower().find(keyword) >= 0

def get_created_date(created):
    mask = "%Y-%m-%dT%H:%M:%S.%f%z"
    return dt.datetime.strptime(created,mask)

def dailyreport(jira, issue):
    comments = jira.comments(issue)
    if len(comments) == 0:
        return ''
    comments_filtered = list(filter(lambda x: comment_clause(x.created,x.body), comments))
    if len(comments_filtered) == 0:
        return ''
    res = ''
    for comment in comments_filtered:
        created_date = get_created_date(comment.created)
        res = res + created_date.strftime("%Y-%m-%d") +"\n" + comment.body + "\n"
    res = res.replace('#dailyreport','')
    res = res.replace('\r','\n')
    res = res.strip('\n')
    res = res.replace('\n\n','\n')
    res = res.replace('\n\n','\n')
    return res