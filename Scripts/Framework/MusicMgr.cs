using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.Events;

public class MusicMgr : SingletonBase<MusicMgr>
{
    // ???????ִ?С
    public float MusicVolume
    {
        get => BGM.volume;
        set => BGM.volume = value;
    }
    // ??Ч??С
    public float SoundVolume
    {
        get => soundVolume;
        set
        {
            soundVolume = value;
            foreach (GameObject sound in soundList)
                sound.GetComponent<AudioSource>().volume = value;
        }
    }
    private float soundVolume;

    // ???????ֶ???
    private AudioSource BGM;
    // ??Ч?б?
    private List<GameObject> soundList;
    // ??Ч?ڻ?????????????
    private const string soundPool = "__Sound";

    public MusicMgr()
    {
        // ???????????ֶ???
        GameObject obj = new GameObject("BGM");
        BGM = obj.AddComponent<AudioSource>();
        BGM.loop = true;

        // ??????Ч?б?
        soundList = new List<GameObject>();

        // ??????Ч?????Ƿ񲥷??? ?򻺴??ع黹??Ч????
        MonoMgr.Instance.AddUpdateListener(() =>
        {
            for (int i = soundList.Count - 1; i >= 0; i--)
            {
                if (!soundList[i].GetComponent<AudioSource>().isPlaying)
                {
                    PoolMgr.Instance.Store(soundPool, soundList[i]);
                    soundList.RemoveAt(i);
                }
            }
        });
        soundVolume = 0.5f;
    }

    #region ???????ַ???
    /// <summary>
    /// ???ű???????
    /// </summary>
    /// <param name="resPath">??????????</param>
    public void PlayMusic(string resPath)
    {
        ResMgr.Instance.LoadAsync<AudioClip>(resPath, (clip) =>
        {
            BGM.clip = clip;
            BGM.Play();
        });
    }
    /// <summary>
    /// ֹͣ???ű???????
    /// </summary>
    public void StopMusic() => BGM.Stop();
    /// <summary>
    /// ??ͣ???ű???????
    /// </summary>
    public void PauseMusic() => BGM.Pause();
    #endregion

    #region ??Ч????
    /// <summary>
    /// ??????Ч
    /// </summary>
    /// <param name="resPath">??Ч??Դ·??</param>
    public void PlaySound(string resPath, bool isLoop = false, UnityAction<GameObject> callback = null)
    {
        ResMgr.Instance.LoadAsync<AudioClip>(resPath, (clip) =>
        {
            // ?ӻ??????м??ض???
            GameObject obj = PoolMgr.Instance.Fetch(soundPool, true); 
            AudioSource sound = obj.GetComponent<AudioSource>();
            if (sound == null)
                sound = obj.AddComponent<AudioSource>();
            // ?????µ???Ч
            sound.clip = clip;
            sound.loop = isLoop;
            sound.volume = soundVolume;
            sound.Play();
            soundList.Add(obj);
            // ????Ч??????????
            if (callback != null)
                callback(obj);
        });
    }
    public void PlaySound(string resPath, Vector3 position, bool isLoop = false, UnityAction<GameObject> callback = null)
    {
        callback += (obj) => obj.transform.position = position;
        PlaySound(resPath, isLoop, callback);
    }
    /// <summary>
    /// ֹͣ??????Ч ?򻺴??ع黹??Ч????
    /// </summary>
    /// <param name="sound">Ҫֹͣ???ŵ???Ч????</param>
    public void StopSound(GameObject sound)
    {
        if (soundList.Contains(sound))
        {
            sound.GetComponent<AudioSource>().Stop();
            soundList.Remove(sound);
            PoolMgr.Instance.Store(soundPool, sound);
        }
    }
    #endregion
}
