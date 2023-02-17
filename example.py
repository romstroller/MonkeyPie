

#!/usr/bin/env python

# from: https://www.mediamonkey.com/forum/viewtopic.php?t=67253

''' Get unique list of all MP3 songs that have been rated
    - Ignore songs with multiple entries, and ignore the rating
    Get list of all FLACs
    Do a Levenshtein fuzzy compare of the "Artist + SongTitle"
    (http://code.google.com/p/pylevenshtein/)
    For any matches, copy the Rating from the MP3 to the matching FLAC
    Update the database and tags
'''
import win32com.client
import sys
import time
import Levenshtein
import codecs

"""
def comp(mp3s, flacs):
    matches = {}
    mp3match = 0
    flacmatch = 0

    # Log the perfect matches
    f100=codecs.open("mp3-flac_matches.100", "w", "utf-8")
    # Log the near matches to look at later, so I can clean the names
    f90=codecs.open("mp3-flac_matches.90", "w", "utf-8")
    # Log any MP3s that don't have a matching FLAC, to see what I'm missing
    nomatch=codecs.open("mp3-flac_matches.no", "w", "utf-8")

    print "Comparing MP3s to FLACs"
    for mp3 in mp3s.keys():
        foundmatch=False
        for flac in flacs:
            ratio = int(Levenshtein.ratio(mp3, flac) * 100)
            if( ratio == 100 ):
                f100.write('%3d   "%s"  ~  "%s"\n' % (ratio, mp3, flac))
                foundmatch=True
                flacmatch += 1
                sys.stdout.write("!")
                matches[mp3] = mp3s[mp3]    # Add match and rating to matches{}
            if( ratio < 100 and ratio >= 90):
                # Log the near matches, which are probably only a character or two different
                f90.write('%3d   "%s"  ~  "%s"\n' % (ratio, mp3, flac))
        if(foundmatch == False):
            nomatch.write(mp3 + "\n")
        else:
            mp3match += 1

    sys.stdout.write("\n")
    f100.close()
    f90.close()
    nomatch.close()

    print"Matched %d MP3 ratings to %d FLACs" % (mp3match, flacmatch)
    return matches


def get_rated_mp3s():
    mp3s={}

    print "Getting MP3s"
    SDB = win32com.client.Dispatch('SongsDB.SDBApplication')
    SDB.ShutdownAfterDisconnect = False
    seltracks = SDB.Database.QuerySongs("Songs.Rating > 0 AND Songs.SongPath LIKE '%.mp3'")
    while not seltracks.EOF:
        trk = seltracks.Item
        key = trk.ArtistName + "~" + trk.Title
        if( mp3s.has_key(key) ):
            if( trk.Rating > mp3s[key] ):
                mp3s[key] = trk.Rating
        else:
            mp3s[key] = trk.Rating
        seltracks.Next()
    print "Found %d rated MP3s" % len(mp3s)
    SDB = None

    return mp3s
"""

def get_flacs():
    flacs=[]

    print "Getting FLACs"
    SDB = win32com.client.Dispatch('SongsDB.SDBApplication')
    SDB.ShutdownAfterDisconnect = False
    seltracks = SDB.Database.QuerySongs("Songs.SongPath LIKE '%.flac'")
    while not seltracks.EOF:
        trk = seltracks.Item
        flacs.append(trk.ArtistName + "~" + trk.Title)
        seltracks.Next()
    print "Found %d FLACs" % len(flacs)
    SDB = None

    flacs.sort()
    return flacs

""" 

def to_unicode_or_bust(obj, encoding='utf-8'):
    ''' Convert object to Unicode. Necessary because of all the
        different possible encodings of entries in MB and MM '''

    if isinstance(obj, basestring):
        if not isinstance(obj, unicode):
            obj = unicode(obj, encoding)
    return obj


def process_matches(matches):
    ''' Update the DB and tags of any matching MP3s and FLACs '''

    keys = matches.keys()
    keys.sort()
    for key in keys:
        rating = matches[key]
        artist, song = key.split("~")
        updatetags(artist, song, rating)
    return


def process_now_playing(mp3s):
    print "Updating NowPlaying with all MP3s that have a raing of 0"

    keys = mp3s.keys()
    keys.sort()
    for key in keys:
        rating = mp3s[key]
        if( rating == 0 ):
            artist, song = key.split("~")
            update_now_playing(artist, song, rating)
    return


def update_now_playing(artist, song, rating):
    # Quote the Artist and Song, which might include ' or "
    # If they contain both, we're screwed, so log and don't proceed
    if("'" in artist and '"' in artist):
        print "Can't quote this artist for QuerySongs: %s" % artist
        log.write("Can't quote this artist for QuerySongs: %s\n" % artist)
        return
    if("'" in song and '"' in song):
        print "Can't quote this song for QuerySongs: %s" % song
        log.write("Can't quote this song for QuerySongs: %s\n" % song)
        return

    if("'" in artist):
        qartist = '"' + artist + '"'
    elif('"' in artist):
        qartist = "'" + artist + "'"
    else:
        qartist = "'" + artist + "'"

    if("'" in song):
        qsong = '"' + song + '"'
    elif('"' in song):
        qsong = "'" + song + "'"
    else:
        qsong = "'" + song + "'"

    SDB = win32com.client.Dispatch('SongsDB.SDBApplication')
    SDB.ShutdownAfterDisconnect = False

    query = "Songs.Artist=" + qartist + " AND Songs.SongTitle=" + qsong + " AND Songs.SongPath LIKE '%.mp3'"
    seltracks = SDB.Database.QuerySongs(query)
    # Add the songs to the NowPlaying list, to make an easy-to-browse list
    # TO-DO: Clear the NowPlaying list before adding all these tracks
    while not seltracks.EOF:
        trk = seltracks.Item
        SDB.Player.PlaylistAddTrack(trk)
        seltracks.Next()

    return


def updatetags(artist, song, rating):
    ''' Update the Rating tag in all the FLAC(s) that match the Artist & Song '''

    log=codecs.open("updatetags.log", "a", "utf-8")

    # Quote the Artist and Song, which might include ' or "
    # If they contain both, we're screwed, so log and don't proceed
    if("'" in artist and '"' in artist):
        print "Can't quote this artist for QuerySongs: %s" % artist
        log.write("Can't quote this artist for QuerySongs: %s\n" % artist)
        return
    if("'" in song and '"' in song):
        print "Can't quote this song for QuerySongs: %s" % song
        log.write("Can't quote this song for QuerySongs: %s\n" % song)
        return

    if("'" in artist):
        qartist = '"' + artist + '"'
    elif('"' in artist):
        qartist = "'" + artist + "'"
    else:
        qartist = "'" + artist + "'"

    if("'" in song):
        qsong = '"' + song + '"'
    elif('"' in song):
        qsong = "'" + song + "'"
    else:
        qsong = "'" + song + "'"

    SDB = win32com.client.Dispatch('SongsDB.SDBApplication')
    SDB.ShutdownAfterDisconnect = False

    query = "Songs.Artist=" + qartist + " AND Songs.SongTitle=" + qsong + " AND Songs.SongPath LIKE '%.flac'"
    seltracks = SDB.Database.QuerySongs(query)
    while not seltracks.EOF:
        trk = seltracks.Item
        print "Setting FLAC: %s ~ %s ~ rating: %d" % (qartist, qsong, rating)
        log.write("Setting FLAC: %s ~ %s ~ rating: %d\n" % (qartist, qsong, rating))
        trk.Rating = rating
        trk.WriteTags()
        trk.UpdateDB()
        seltracks.Next()

    log.close()
    return


def main():
    mp3s  = get_rated_mp3s()
    process_now_playing(mp3s)
    flacs = get_flacs()
    matches = comp(mp3s,flacs)
    process_matches(matches)
    print "Done."


if __name__ == '__main__':
    main()
"""