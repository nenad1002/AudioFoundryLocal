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

    // ----------------------------------------------------------------
    // COM control interface for the streaming transcriber, using IAudioStream
    // ----------------------------------------------------------------
    [ComVisible(true)]
    [Guid("AABBCCDD-7777-8888-9999-AAAABBBBCCCC")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAudioStreamingTranscriber
    {
        [DispId(1)] void Initialize(IAudioConfig config);
        [DispId(2)] void RegisterCallback(ITranscriptionCallback callback);
        [DispId(3)]
        void Start(IStream audioStream);
        [DispId(4)] void Stop();
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
        private IStream? _stream;
        private CancellationTokenSource? _cts;
        private Task? _processingTask;
        // TODO: revise buffer size
        private readonly int _bufferSize = 4096;

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

        public void Start(IStream audioStream)
        {
            if (_config == null)
                throw new InvalidOperationException("Initialize must be called first.");

            _stream = audioStream;
            _cts = new CancellationTokenSource();
            _processingTask = Task.Run(() => ProcessStreamLoop(_cts.Token), _cts.Token);
        }

        public void Stop()
        {
            _cts?.Cancel();
            try { _processingTask?.Wait(); } catch (AggregateException ae) when (ae.InnerException is OperationCanceledException) { }
        }

        private void ProcessStreamLoop(CancellationToken token)
        {
            var buffer = new byte[_bufferSize];
            while (!token.IsCancellationRequested)
            {
                int bytesRead;
                IntPtr readPtr = Marshal.AllocCoTaskMem(sizeof(int));
                try
                {
                    _stream!.Read(buffer, _bufferSize, readPtr);
                    bytesRead = Marshal.ReadInt32(readPtr);
                }
                catch (Exception ex)
                {
                    _callback?.OnError(ex.Message);
                    break;
                }
                finally
                {
                    Marshal.FreeCoTaskMem(readPtr);
                }

                if (bytesRead <= 0)
                {
                    Thread.Sleep(10);
                    continue;
                }

                // TODO: here it goes a real call into Whisper.
                string partial = TranscribePartial(buffer, bytesRead);
                _callback?.OnPartialResult(partial);
            }
            _callback?.OnFinalResult("[stream ended]");
        }

        private string TranscribePartial(byte[] chunk, int length)
        {
            return $"[partial {length} bytes]";
        }
    }
}
