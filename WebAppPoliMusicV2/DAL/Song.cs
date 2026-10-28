using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Web;

namespace WebAppPoliMusicV2.DAL
{
    /// <summary>
    /// Data access layer for the [TBL_SONG] table.
    /// </summary>
    public class Song
    {
        private string strConnectionString;

        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }

        /// <summary>
        /// Inserts a new song on [TBL_SONG].
        /// </summary>
        /// <returns>Number of affected rows.</returns>
        public int Save(Entities.Song song)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "INSERT INTO [TBL_SONG]([SONG_NAME], [SONG_PATH], [PLAYS]) " +
                        "VALUES(@SONG_NAME, @SONG_PATH, @PLAYS)";
                    cmd.Parameters.AddWithValue("@SONG_NAME", song.SongName);
                    cmd.Parameters.AddWithValue("@SONG_PATH", song.SongPath);
                    cmd.Parameters.AddWithValue("@PLAYS", song.Plays);
                    cmd.Connection = con;
                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Returns all the songs ordered by id.
        /// </summary>
        public DataTable Read()
        {
            DataTable dtSong = new DataTable();
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand("SELECT [ID_SONG], [SONG_NAME], [SONG_PATH], ISNULL([PLAYS], 0) AS [PLAYS] " +
                    "FROM [TBL_SONG] ORDER BY [ID_SONG]", con))
                {
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(dtSong);
                        return dtSong;
                    }
                }
            }
        }

        /// <summary>
        /// Returns one song by its id.
        /// </summary>
        public DataTable Read(int idSong)
        {
            DataTable dtSong = new DataTable();
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand("SELECT [ID_SONG], [SONG_NAME], [SONG_PATH], ISNULL([PLAYS], 0) AS [PLAYS] " +
                    "FROM [TBL_SONG] WHERE [ID_SONG] = @ID_SONG", con))
                {
                    cmd.Parameters.AddWithValue("@ID_SONG", idSong);
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(dtSong);
                        return dtSong;
                    }
                }
            }
        }

        /// <summary>
        /// Updates an existing song.
        /// </summary>
        /// <returns>Number of affected rows.</returns>
        public int Update(Entities.Song song)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "UPDATE [TBL_SONG] SET " +
                        "[SONG_NAME] = @SONG_NAME " +
                        ", [SONG_PATH] = @SONG_PATH " +
                        ", [PLAYS] = @PLAYS " +
                        "WHERE [ID_SONG] = @ID_SONG";
                    cmd.Parameters.AddWithValue("@ID_SONG", song.ID);
                    cmd.Parameters.AddWithValue("@SONG_NAME", song.SongName);
                    cmd.Parameters.AddWithValue("@SONG_PATH", song.SongPath);
                    cmd.Parameters.AddWithValue("@PLAYS", song.Plays);
                    cmd.Connection = con;
                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Deletes a song. The rows that reference the song are removed first
        /// to respect the foreign keys of the database.
        /// </summary>
        /// <returns>Number of affected rows.</returns>
        public int Delete(int idSong)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "DELETE [TBL_TRACKLIST] WHERE [ID_SONG] = @ID_SONG; " +
                        "DELETE [TBL_PLAYER] WHERE [ID_SONG] = @ID_SONG; " +
                        "DELETE [TBL_PLAYLIST_SONG] WHERE [ID_SONG] = @ID_SONG; " +
                        "DELETE [TBL_SONG] WHERE [ID_SONG] = @ID_SONG;";
                    cmd.Parameters.AddWithValue("@ID_SONG", idSong);
                    cmd.Connection = con;
                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// Increases the plays counter of a song and returns the new value.
        /// </summary>
        public int RegisterPlay(int idSong)
        {
            using (SqlConnection con = new SqlConnection(strConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.CommandText = "UPDATE [TBL_SONG] SET [PLAYS] = ISNULL([PLAYS], 0) + 1 " +
                        "OUTPUT INSERTED.[PLAYS] " +
                        "WHERE [ID_SONG] = @ID_SONG";
                    cmd.Parameters.AddWithValue("@ID_SONG", idSong);
                    cmd.Connection = con;
                    con.Open();
                    object returnValue = cmd.ExecuteScalar();
                    return (returnValue == null) ? 0 : Convert.ToInt32(returnValue);
                }
            }
        }
    }
}
