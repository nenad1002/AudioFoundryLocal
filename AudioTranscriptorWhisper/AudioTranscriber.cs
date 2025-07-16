using Microsoft.VisualBasic;
using NAudio.Wave;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Channels;
using System.Threading.Tasks;

namespace AudioTranscriberLib
{
    // ----------------------------------------------------------------
    // COM audio configuration interface (model-specific)
    // ----------------------------------------------------------------
    [ComVisible(true)]
    [Guid("11223344-5566-7788-99AA-BBCCDDEEFF00")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAudioConfig
    {
        [DispId(1)] string ModelAlias { get; }
        [DispId(2)] string Language { get; }
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
    }

    // ----------------------------------------------------------------
    // COM callback interface for raw audio data
    // ----------------------------------------------------------------
    [ComVisible(true)]
    [Guid("ABCDEF00-1111-2222-3333-444455556677")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAudioDataCallback
    {
        [DispId(1)]
        void OnAudioData(
            [MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_UI1)]
            byte[] pcmChunk);
    }

    // ----------------------------------------------------------------
    // COM audio stream interface
    // ----------------------------------------------------------------
    [ComVisible(true)]
    [Guid("A1B2C3D4-1234-5678-9ABC-DEF012345678")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAudioStream
    {
        [DispId(1)] int SampleRate { get; }
        [DispId(2)] int ChunkSize { get; }
        [DispId(3)] int DeviceIndex { get; }
        [DispId(4)]
        void RegisterCallback(IAudioDataCallback callback);
    }

    // ----------------------------------------------------------------
    // COM audio stream
    // ----------------------------------------------------------------
    [ComVisible(true)]
    [Guid("A1B2C3D4-1234-5678-9ABC-DEF012345679")]
    [ProgId("AudioTranscriber.AudioStream")]
    [ClassInterface(ClassInterfaceType.None)]
    public class AudioStream : IAudioStream
    {
        private IAudioDataCallback? _sink;
        private readonly WaveInEvent _capture;

        public AudioStream(int sampleRate, int chunkSize, int deviceIndex)
        {
            SampleRate = sampleRate;
            ChunkSize = chunkSize;
            DeviceIndex = deviceIndex;

            _capture = new WaveInEvent
            {
                DeviceNumber = DeviceIndex,
                WaveFormat = new WaveFormat(SampleRate, 16, 1),
                BufferMilliseconds = (int)(1000.0 * ChunkSize / SampleRate)
            };
            _capture.DataAvailable += (s, e) =>
            {
                _sink?.OnAudioData(e.Buffer);
            };
        }

        public int SampleRate { get; }
        public int ChunkSize { get; }
        public int DeviceIndex { get; }

        public void RegisterCallback(IAudioDataCallback callback)
        {
            _sink = callback;
            _capture.StartRecording();
        }
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
    // COM control interface for the transcriber
    // ----------------------------------------------------------------
    [ComVisible(true)]
    [Guid("AABBCCDD-7777-8888-9999-AAAABBBBCCCC")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAudioTranscriber
    {
        [DispId(1)] void Initialize(IAudioConfig config);
        [DispId(2)] void RegisterCallback(ITranscriptionCallback callback);
        [DispId(3)] void Start(IAudioStream stream);
        [DispId(4)] void Stop();
    }

    // ----------------------------------------------------------------
    // COM coclass implementing the transcriber and audio-data sink
    // ----------------------------------------------------------------
    [ComVisible(true)]
    [Guid("DDCCBBAA-CCCC-BBBB-AAAA-999988887777")]
    [ProgId("AudioTranscriber.Component")]
    [ClassInterface(ClassInterfaceType.None)]
    public class AudioTranscriber : IAudioTranscriber, IAudioDataCallback
    {
        private ITranscriptionCallback? _callback;
        private IAudioConfig? _config;
        private IAudioStream? _stream;
        private Channel<byte[]>? _channel;
        private CancellationTokenSource? _cts;
        private Task? _processingTask;

        // use the constructor as the input entry
        public AudioTranscriber(IAudioConfig config)
        {
            this.Initialize(config);
        }

        public void Initialize(IAudioConfig config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
        }

        public void RegisterCallback(ITranscriptionCallback callback)
        {
            _callback = callback;
        }

        public void Start(IAudioStream stream)
        {
            if (_config == null)
                throw new InvalidOperationException("Call Initialize(config) before Start(stream).");

            _stream = stream ?? throw new ArgumentNullException(nameof(stream));

            _channel = Channel.CreateUnbounded<byte[]>(new UnboundedChannelOptions
            {
                SingleReader = true,
                SingleWriter = false
            });
            _cts = new CancellationTokenSource();

            // Start the loop.
            _processingTask = ProcessLoopAsync(_cts.Token);

            // Register for automatic audio data events
            _stream.RegisterCallback(this);
        }

        public void Stop()
        {
            if (_stream is AudioStream audioStream)
            {
                // Stop recording.
                audioStream?.GetType().GetMethod("StopRecording")?.Invoke(audioStream, null);
            }

            // Cancel transcription loop
            if (_cts == null) return;
            _cts.Cancel();
            try
            {
                _processingTask?.Wait();
            }
            catch (AggregateException ae) when (ae.InnerException is OperationCanceledException)
            {
                // expected
            }
        }

        // Receives raw PCM from IAudioStream
        public void OnAudioData(byte[] pcmChunk)
        {
            if (_cts?.IsCancellationRequested == true) return;

            // Write it on the channel to be consumed by the processing loop and fed into a model.
            _channel?.Writer.TryWrite(pcmChunk);
        }

        private async Task ProcessLoopAsync(CancellationToken token)
        {
            try
            {
                await foreach (var chunk in _channel!.Reader.ReadAllAsync(token))
                {
                    // Here feed to Whisper or FoundryLocalManager.
                    string partial = await TranscribePartialAsync(chunk, token);
                    _callback?.OnPartialResult(partial);
                }
            }
            catch (OperationCanceledException)
            {
                // normal shutdown
            }
            catch (Exception ex)
            {
                _callback?.OnError(ex.Message);
            }
        }

        // Replace with actual code call into Whisper
        private Task<string> TranscribePartialAsync(byte[] chunk, CancellationToken token)
        {
            return Task.FromResult($"[partial {chunk.Length} bytes @ {_stream?.SampleRate}Hz]");
        }
    }
}
