---
layout: post
title:  "testing code markups and screenshots"
date:   2017-08-19 12:12:01 -0600
categories: jekyll testing
---

normal text

{% highlight ruby %}
def send_email(to, msg, new_template):
        
    data.reverse()    
    
    for i in range(0, 6):
        data.pop()
    #    print(email_list)
           
    try:#outlook.com keeps sending the FIRST(and only the FIRST) email.... how to check?
        data.reverse()
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.login(gmail_user, gmail_pword)
        server.sendmail(sent_from, to, msg.as_string())
        
        print("Email Sent To: ", new_template.name)
        print("@: ", to)            
        print("Invoice Numbers: ", new_template.invoice_num)
        print("TOTAL: ", new_template.total, "---------------------------")            
        server.quit()
        
        if(data[0] == 'None'):
            print('END OF LIST')
        else:            
            build_email(data)
    except Exception as e:
        print(e)
        print('Email Failed to Send to: ', new_template.name)
        print("@: ", to)
        print("Invoice Numbers: ", new_template.invoice_num)
data = list_invoices()
#print(data)
{% endhighlight %}

screenshot here: 
[My helpful screenshot]({{ blackandbluewater.com }}/assets/testing-screen-shot.png)

hyperlink: [jekyll how to](https://jekyllrb.com/docs/posts/)
