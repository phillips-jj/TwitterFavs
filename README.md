# TwitterFavs
Twitter-Excel integration sample

This project highlights an integration between the Twitter API and Excel.  Specifically, it uses the TweetSharp wrapper for the Twitter API and the interop services for Microsoft Excel to download the "liked" (or "favorite") tweets from an account, and write them into an Excel file on disk.  It also highlights multithreaded operation, by running the "hard work" in a background thread, while posting updates to the main/UI thread (which is best practice for many lanugages and frameworks).  Code is written in C#.
