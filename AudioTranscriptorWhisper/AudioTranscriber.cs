using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace AudioTranscriberLib
{
    // ----------------------------------------------------------------
    // COM audio configuration interface
    // ----------------------------------------------------------------
    [ComVisible(true)]
    [Guid("11223344-5566-7788-99AA-BBCCDDEEFF00")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAudioConfig
    {
        [DispId(0)] int Version { get; }
        [DispId(1)] string ModelAlias { get; }
        [DispId(2)] string Language { get; }
        [DispId(3)] int SampleRate { get; }
        [DispId(4)] int Channels { get; }
        [DispId(5)] int BitsPerSample { get; }
    }

    // ----------------------------------------------------------------
    // COM audio configuration
    // ----------------------------------------------------------------
    [ComVisible(true)]
    [Guid("00FFEEDD-CCBB-AA99-8877-665544332211")]
    [ProgId("AudioTranscriber.AudioConfig")]
    [ClassInterface(ClassInterfaceType.None)]
    public class AudioConfig : IAudioConfig
    {
        public AudioConfig(string modelAlias, string language)
        {
            ModelAlias = modelAlias;
            Language = language;
        }

        public string ModelAlias { get; }
        public string Language { get; }

        public int SampleRate => throw new NotImplementedException();

        public int Channels => throw new NotImplementedException();

        public int BitsPerSample => throw new NotImplementedException();

        public int Version => 1;
    }

    // ----------------------------------------------------------------
    // COM callback interface for transcription results
    // ----------------------------------------------------------------
    [ComVisible(true)]
    [Guid("F1A2B3C4-1111-2222-3333-444455556666")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITranscriptionCallback
    {
        [DispId(1)] void OnPartialResult(string text);
        [DispId(2)] void OnFinalResult(string text);
        [DispId(3)] void OnError(string message);
    }

    [ComVisible(true)]
    [Guid("A1B2C3D4-1111-2222-3333-444455556666")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAudioPacket
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_UI1)]
        byte[] Data { get; }

        [DispId(2)]
        long Timestamp100ns { get; }

        [DispId(3)]
        int SequenceNumber { get; }
    }

    [ComVisible(true)]
    [Guid("D4C3B2A1-6666-5555-4444-333322221111")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAudioPacketCallback
    {
        [DispId(1)]
        void OnPacket(IAudioPacket packet);

        [DispId(2)]
        void OnError(string message);
    }

    [ComVisible(true)]
    [Guid("E1F2A3B4-7777-8888-9999-AAAA0000BBBB")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAudioStreamingTranscriber
    {
        [DispId(1)] void Initialize(IAudioConfig config);
        [DispId(2)] void RegisterTranscriptionCallback(ITranscriptionCallback cb);
        [DispId(3)] void RegisterAudioCallback(IAudioPacketCallback cb);

        [DispId(4)] void Start();
        [DispId(5)] void Stop();

        [DispId(6)]
        void PushPacket(IAudioPacket packet);
    }

    // ----------------------------------------------------------------
    // COM control interface for the non-streaming transcriber, using IAudioStream
    // ----------------------------------------------------------------
    [ComVisible(true)]
    [Guid("AABBCCDD-7777-8888-9999-AAAABBBBCCCC")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAudioNonStreamingTranscriber
    {
        [DispId(1)] void Initialize(IAudioConfig config);
        [DispId(2)]
        void TranscribeBufferAsync(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_UI1)]
        byte[] audioBuffer,
            ITranscriptionCallback callback
        );
        // Cancel in‑flight buffer transcription.
        [DispId(3)]
        void CancelBufferTranscription();
    }

    // ----------------------------------------------------------------
    // COM coclass implementing the transcriber
    // ----------------------------------------------------------------
    [ComVisible(true)]
    [Guid("DDCCBBAA-CCCC-BBBB-AAAA-999988887777")]
    [ProgId("AudioTranscriber.Component")]
    [ClassInterface(ClassInterfaceType.None)]
    public class AudioTranscriber : IAudioStreamingTranscriber
    {
        private ITranscriptionCallback? _callback;
        private IAudioConfig? _config;
        private IAudioPacketCallback _audioCb;

        public AudioTranscriber(IAudioConfig config)
        {
            Initialize(config);
        }

        public void Initialize(IAudioConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        public void RegisterCallback(ITranscriptionCallback callback)
        {
            _callback = callback;
        }

        private string TranscribePartial(byte[] chunk, int length)
        {
            return $"[partial {length} bytes]";
        }

        public void RegisterTranscriptionCallback(ITranscriptionCallback cb)
        {
            throw new NotImplementedException();
        }

        public void RegisterAudioCallback(IAudioPacketCallback cb)
        {
            throw new NotImplementedException();
        }

        public void Start()
        {
            throw new NotImplementedException();
        }

        public void PushPacket(IAudioPacket packet)
        {
            try
            {
                // Buffer, chunk as appropriate and call into Whisper
                _audioCb.OnPacket(packet);

                // Should probably be within OnPacket impl.
                _callback.OnPartialResult($"Packet {packet.SequenceNumber} received");

                if (packet.SequenceNumber % 5 == 0)
                {
                    _callback.OnFinalResult($"Completed up to packet {packet.SequenceNumber}");
                }
            }
            catch (Exception ex)
            {
                _audioCb.OnError(ex.Message);
                _callback.OnError(ex.Message);
            }
        }

        public void Stop()
        {
            throw new NotImplementedException();
        }
    }
}
