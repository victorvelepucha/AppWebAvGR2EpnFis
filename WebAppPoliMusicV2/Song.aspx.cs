using System;
using System.Collections.Generic;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV2
{
    public partial class Song : System.Web.UI.Page
    {
        string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                if (!LoadUserInformation())
                {
                    return;
                }
                LoadSongs();
            }
        }

        /// <summary>
        /// Shows the name and the photo of the authenticated user.
        /// </summary>
        /// <returns>False when the user is not authenticated and the page was redirected.</returns>
        private bool LoadUserInformation()
        {
            Entities.User user = Session[Util.Constants.USER] as Entities.User;
            if (user == null)
            {
                // The session is no longer valid, the user has to authenticate again.
                FormsAuthentication.SignOut();
                Response.Redirect("Login.aspx", true);
                return false;
            }
            lblUserName.Text = user.UserName;

            string photo = user.Photo;
            if (string.IsNullOrEmpty(photo))
            {
                photo = "~/UserImages/default.jpg";
            }
            imgUserPhoto.ImageUrl = ResolveUrl(photo);
            imgUserPhoto.AlternateText = user.UserName;
            return true;
        }

        /// <summary>
        /// Loads the songs from the BLL and binds them to the repeater.
        /// </summary>
        private void LoadSongs()
        {
            try
            {
                List<Entities.Song> songs = new BLL.Song(strConnString).GetSongs();

                rptSongs.DataSource = songs;
                rptSongs.DataBind();

                lblCount.Text = songs.Count.ToString();
                pnlSongs.Visible = songs.Count > 0;
                pnlEmpty.Visible = songs.Count == 0;
            }
            catch (Exception ex)
            {
                pnlSongs.Visible = false;
                pnlEmpty.Visible = false;
                lblMessage.Text = "It was not possible to load the song list: " + ex.Message;
                lblMessage.Visible = true;
            }
        }

        /// <summary>
        /// Returns the value encoded as HTML so it can be printed safely inside the table.
        /// </summary>
        protected string Text(object value)
        {
            return Server.HtmlEncode(Convert.ToString(value));
        }

        /// <summary>
        /// Converts the path stored in the database (~/songFiles/song.mp3) into
        /// a path that the browser can request.
        /// </summary>
        protected string GetSongUrl(object songPath)
        {
            string path = Convert.ToString(songPath);
            if (string.IsNullOrEmpty(path))
            {
                return string.Empty;
            }
            path = path.Trim().Replace("\\", "/");
            if (path.StartsWith("http://") || path.StartsWith("https://"))
            {
                return path;
            }
            if (path.StartsWith("~/"))
            {
                path = path.Substring(2);
            }
            else if (path.StartsWith("/"))
            {
                path = path.Substring(1);
            }
            return ResolveUrl("~/" + path);
        }
    }
}
