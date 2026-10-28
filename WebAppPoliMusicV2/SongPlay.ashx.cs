using System;
using System.Configuration;
using System.Web;

namespace WebAppPoliMusicV2
{
    /// <summary>
    /// Increases the plays counter of a song without reloading the page.
    /// It is called from Song.aspx when the user starts a reproduction.
    /// </summary>
    public class SongPlay : IHttpHandler
    {
        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "application/json";
            context.Response.Cache.SetCacheability(HttpCacheability.NoCache);

            if (!context.User.Identity.IsAuthenticated)
            {
                context.Response.StatusCode = 401;
                context.Response.Write("{\"error\":\"unauthorized\"}");
                return;
            }

            int idSong;
            if (!int.TryParse(context.Request.QueryString["id"], out idSong) || idSong <= 0)
            {
                context.Response.StatusCode = 400;
                context.Response.Write("{\"error\":\"invalid song id\"}");
                return;
            }

            try
            {
                string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;
                int plays = new BLL.Song(strConnString).RegisterPlay(idSong);
                context.Response.Write("{\"id\":" + idSong + ",\"plays\":" + plays + "}");
            }
            catch (Exception ex)
            {
                context.Response.StatusCode = 500;
                string message = ex.Message.Replace("\\", "/").Replace("\"", "'").Replace("\r", " ").Replace("\n", " ");
                context.Response.Write("{\"error\":\"" + message + "\"}");
            }
        }

        public bool IsReusable
        {
            get { return false; }
        }
    }
}
