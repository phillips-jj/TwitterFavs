using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Xml;
using System.Windows.Forms;
using TweetSharp;

namespace TwitterFavs
{
    public partial class frmMain : Form
    {
        //these are values from the Twitter API page, specific for your account
        private string mConsumerKey;
        private string mConsumerSecret;
        private string mAccessToken;
        private string mAccessTokenSecret;

        private BackgroundWorker mWorkerThread;
        private long mTweetsSaved;

        
        public frmMain()
        {
            InitializeComponent();
        }


        private void ExportTweetsProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            mTweetsSaved = (long)e.UserState;
            stripLabel.Text = "Saved " + mTweetsSaved + " tweets.";
        }

        private void ExportTweetsCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            stripLabel.Text = "Done - saved " + mTweetsSaved + " tweets.";
            MessageBox.Show("Export Completed Successfully");
        }


        private void ExportTweets(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker callback = (BackgroundWorker)sender; //used to report progress back to the UI thread
            
            ExcelTweetWriter tweetWriter = new ExcelTweetWriter((string)e.Argument);

            //NB - the credentials that are pre-packaged are invalid - you need to use your own, valid credentials
            TwitterService serv = new TwitterService(mConsumerKey, mConsumerSecret);
            serv.AuthenticateWith(mAccessToken, mAccessTokenSecret);

            ListFavoriteTweetsOptions favopts = new ListFavoriteTweetsOptions();
            favopts.Count = 200;    //read max 200 tweets per iteration

            List<TwitterStatus> tweets;

            long l = 1;
            bool CheckForMore = true;
            while (CheckForMore)
            {
                tweets = serv.ListFavoriteTweets(favopts).ToList();
                int numTweets = tweets.Count;
                if (numTweets > 0)
                {
                    foreach (var tweet in tweets)
                    {
                        tweetWriter.Write(tweet.CreatedDate, tweet.Id, tweet.User.ScreenName, tweet.User.Name, tweet.Text);
                        ++l;
                    }

                    favopts.MaxId = (tweets[numTweets - 1].Id) - 1; //tell the Twitter API where (by result count)to resume fetching tweets

                    callback.ReportProgress(0, l);  //unknown percentage, but we can report the number of tweets we've read thus far
                }
                else
                {
                    CheckForMore = false;
                }
            }

            tweetWriter.Close();
            tweetWriter = null;
        }


        private void btnGo_Click(object sender, EventArgs e)
        {
            btnGo.Enabled = false;
            txtFile.Enabled = false;
            btnFile.Enabled = false;

            stripLabel.Text = "Beginning export...";

            mWorkerThread = new BackgroundWorker();
            mWorkerThread.DoWork += new DoWorkEventHandler(ExportTweets);
            mWorkerThread.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ExportTweetsCompleted);
            mWorkerThread.WorkerReportsProgress = true;
            mWorkerThread.ProgressChanged += new ProgressChangedEventHandler(ExportTweetsProgressChanged);
            mWorkerThread.RunWorkerAsync(txtFile.Text);
        }


        private bool LoadCredentials()
        {
            bool bSuccess = false;
            bool bExceptionCaught = false;
            const string sFilename = "twitter-credentials.xml";
            const string sErrorMessageStart = "Unable to open the file " + sFilename + ". ";

            //load XML file containing keys/secrets/tokens, and read them
            try
            {
                XmlReader xmlCredentials = XmlReader.Create(sFilename);
                bSuccess = xmlCredentials.ReadToFollowing("credentials");
                if (bSuccess)
                {
                    mConsumerKey = xmlCredentials.GetAttribute("ConsumerKey");
                    if ((mConsumerKey == null) || (mConsumerKey.Length == 0))
                    {
                        bSuccess = false;
                    }

                    mConsumerSecret = xmlCredentials.GetAttribute("ConsumerSecret");
                    if ((mConsumerSecret == null) || (mConsumerSecret.Length == 0))
                    {
                        bSuccess = false;
                    }

                    mAccessToken = xmlCredentials.GetAttribute("AccessToken");
                    if ((mAccessToken == null) || (mAccessToken.Length == 0))
                    {
                        bSuccess = false;
                    }

                    mAccessTokenSecret = xmlCredentials.GetAttribute("AccessTokenSecret");
                    if ((mAccessTokenSecret == null) || (mAccessTokenSecret.Length == 0))
                    {
                        bSuccess = false;
                    }
                }

                xmlCredentials.Close();
            }
            catch (System.Security.SecurityException /*e*/)
            {
                bExceptionCaught = true;
                MessageBox.Show(sErrorMessageStart + "Check the security settings of the file.");
            }
            catch (System.IO.FileNotFoundException /*e*/)
            {
                bExceptionCaught = true;
                MessageBox.Show(sErrorMessageStart + "Ensure that the file is in the same directory as this executable.");
            }
            catch (Exception /*e*/)
            {
                bExceptionCaught = true;
                MessageBox.Show(sErrorMessageStart + "Ensure that the file is not corrupt.");
            }

            if (bExceptionCaught)
            {
                bSuccess = false;
            }

            if (!bSuccess && !bExceptionCaught)
            {
                MessageBox.Show("One or more of the required attributes could not be read from " + sFilename + ".  Ensure the file formatting is correct, and the data is present.");
            }

            return bSuccess;
        }


        private void frmMain_Load(object sender, EventArgs e)
        {
            bool bCredentialsLoaded = LoadCredentials();
            if (bCredentialsLoaded)
            {
                txtFile.ReadOnly = true;
                stripLabel.Text = "Ready...";

                SetFileUI(System.IO.Directory.GetCurrentDirectory() + "\\twitter-favs.xlsx");
            }
            else
            {
                txtFile.Enabled = false;
                btnFile.Enabled = false;
                btnGo.Enabled = false;
            }

        }


        private void SetFileUI (string fileName)
        {
            txtFile.Text = fileName;

            txtFile.SelectionStart = txtFile.TextLength;
            txtFile.ScrollToCaret();
        }


        private void btnFile_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "Excel files (*.xlsx)|*.xlsx";
            dlg.FilterIndex = 1;
            dlg.RestoreDirectory = true;
            dlg.InitialDirectory = System.IO.Directory.GetCurrentDirectory();
            dlg.FileName = "twitter-favs.xlsx";

            DialogResult dlgres = dlg.ShowDialog();
            if (dlgres == System.Windows.Forms.DialogResult.OK)
            {
                SetFileUI(dlg.FileName);
            }
        }
    }
}
