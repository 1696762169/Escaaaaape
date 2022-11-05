using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.Events;

public class MusicMgr : SingletonBase<MusicMgr>
{
    // 背景音乐大小
    public float MusicVolume
    {
        get => BGM.volume;
        set => BGM.volume = value;
    }
    // 音效大小
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

    // 背景音乐对象
    private AudioSource BGM;
    // 音效列表
    private List<GameObject> soundList;
    // 音效在缓存池里的名称
    private const string soundPool = "__Sound";

    public MusicMgr()
    {
        // 创建背景音乐对象
        GameObject obj = new GameObject("BGM");
        BGM = obj.AddComponent<AudioSource>();
        BGM.loop = true;

        // 创建音效列表
        soundList = new List<GameObject>();

        // 检测音效对象是否播放完 向缓存池归还音效对象
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

    #region 背景音乐方法
    /// <summary>
    /// 播放背景音乐
    /// </summary>
    /// <param name="resPath">背景音乐名</param>
    public void PlayMusic(string resPath)
    {
        ResMgr.Instance.LoadAsync<AudioClip>(resPath, (clip) =>
        {
            BGM.clip = clip;
            BGM.Play();
        });
    }
    /// <summary>
    /// 停止播放背景音乐
    /// </summary>
    public void StopMusic() => BGM.Stop();
    /// <summary>
    /// 暂停播放背景音乐
    /// </summary>
    public void PauseMusic() => BGM.Pause();
    #endregion

    #region 音效方法
    /// <summary>
    /// 播放音效
    /// </summary>
    /// <param name="resPath">音效资源路径</param>
    public void PlaySound(string resPath, bool isLoop = false, UnityAction<GameObject> callback = null)
    {
        ResMgr.Instance.LoadAsync<AudioClip>(resPath, (clip) =>
        {
            // 从缓存池中加载对象
            GameObject obj = PoolMgr.Instance.Fetch(soundPool, true); 
            AudioSource sound = obj.GetComponent<AudioSource>();
            if (sound == null)
                sound = obj.AddComponent<AudioSource>();
            // 设置新的音效
            sound.clip = clip;
            sound.loop = isLoop;
            sound.volume = soundVolume;
            sound.Play();
            soundList.Add(obj);
            // 对音效的其它处理
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
    /// 停止播放音效 向缓存池归还音效组件
    /// </summary>
    /// <param name="sound">要停止播放的音效对象</param>
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
