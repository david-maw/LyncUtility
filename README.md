# LyncUtility

This is a simple WPF program to set a status in Lync (Skype for business) for a user defined period, then revert to a user specified status (or an automatic one provided by Lync). It is written in C# and compiled using VS2015 (though there are no features in it requiring that release).

You can find more information on it at http://www.codeproject.com/Articles/1115739/LyncUtility-a-Program-to-Manage-Lync-Skype-for-Business.

It uses COM interop to talk to the Lync client on your PC and uses the Lync SDK (which you'll need to download from Microsoft and install if you want to recompile this) to provide the features it needs.

The basic problem it is trying to solve is that I kept wanting to set my status to "away" for a while, then would forget to reset it later so I'd be marked as "away" for a lot longer than I wanted. By allowing you to set a status for a fixed period, knowing it'll be automatically changed later, this program prevents that particular error.
