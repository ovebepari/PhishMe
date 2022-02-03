I started making a plugin for Microsoft Outlook desktop client and for the Web Outlook as well. My client organization receives a lot of spam emails and they were trying to report them to their sysadmin. "Phish alert" or similiar kind of a plugin they wanted, which will forward the current email to their sysadmin, in one click.

![Tux, the Linux mascot](https://raw.githubusercontent.com/ovebepari/ovebepari.github.io/d48d140a0c1a77d2a50757c07884ba169701a07a/post_images/2021/Mar/single_click_PhishMe.png)

However, there are two API endpoints to access and work with user emails.

    MS Office API (Deprecating)
    MS Graph API (Recommended)

I built my addon with MS Office API first. Then I needed to send email threads as attachments which I could't find any documentations within MS Office API, it was doable, but needed to hit MS Graph API endpoint. For some reason, MS Graph API authentication as developer seemed hectic and the whole task made me overwhelmed to say the least.

You can checkout the repository at PhishMe at Github.
