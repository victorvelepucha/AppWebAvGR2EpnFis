using System;
using System.Collections.Generic;
using System.Data;
using System.Web;

namespace WebAppPoliMusicV2.BLL
{
    /// <summary>
    /// Business layer for the songs. It validates the data and delegates the
    /// persistence to the DAL.
    /// </summary>
    public class Song
    {
        private string strConnectionString;

        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }

        public int Save(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Save(song);
        }

        public DataTable Read()
        {
            return new DAL.Song(strConnectionString).Read();
        }

        public DataTable Read(int idSong)
        {
            return new DAL.Song(strConnectionString).Read(idSong);
        }

        public int Update(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Update(song);
        }

        public int Delete(int idSong)
        {
            return new DAL.Song(strConnectionString).Delete(idSong);
        }

        public int RegisterPlay(int idSong)
        {
            return new DAL.Song(strConnectionString).RegisterPlay(idSong);
        }

        /// <summary>
        /// Returns the song list as a list of entities, ready to be bound
        /// to the web controls.
        /// </summary>
        public List<Entities.Song> GetSongs()
        {
            List<Entities.Song> songs = new List<Entities.Song>();
            DataTable dtSong = Read();
            foreach (DataRow row in dtSong.Rows)
            {
                Entities.Song song = new Entities.Song();
                song.ID = Convert.ToInt32(row["ID_SONG"]);
                song.SongName = row["SONG_NAME"].ToString();
                song.SongPath = row["SONG_PATH"].ToString();
                song.Plays = Convert.ToInt32(row["PLAYS"]);
                songs.Add(song);
            }
            return songs;
        }

        /// <summary>
        /// Returns one song by its id or null if it does not exist.
        /// </summary>
        public Entities.Song GetSong(int idSong)
        {
            DataTable dtSong = Read(idSong);
            if (dtSong == null || dtSong.Rows.Count == 0)
            {
                return null;
            }
            DataRow row = dtSong.Rows[0];
            return new Entities.Song(
                Convert.ToInt32(row["ID_SONG"]),
                row["SONG_NAME"].ToString(),
                row["SONG_PATH"].ToString(),
                Convert.ToInt32(row["PLAYS"]));
        }

        /// <summary>
        /// Validates the required fields of the song before saving it.
        /// </summary>
        /// <returns>An empty string if the song is valid or the error message.</returns>
        public string ValidationsRequired(Entities.Song song)
        {
            if (song.SongName == null || song.SongName.Trim() == string.Empty)
            {
                return "The song name is required!";
            }
            if (song.SongPath == null || song.SongPath.Trim() == string.Empty)
            {
                return "The song path is required!";
            }
            if (song.Plays < 0)
            {
                return "The plays counter cannot be negative!";
            }
            return string.Empty;
        }
    }
}
