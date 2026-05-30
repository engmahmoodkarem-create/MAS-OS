import { basename } from "path";
import {
  memo,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import StyledVoiceRecorder from "components/apps/VoiceRecorder/StyledVoiceRecorder";
import { type ComponentProcessProps } from "components/system/Apps/RenderComponent";
import useTitle from "components/system/Window/useTitle";
import { useFileSystem } from "contexts/fileSystem";
import Button from "styles/common/Button";
import { DESKTOP_PATH } from "utils/constants";
import { blobToBuffer } from "utils/functions";

const VoiceRecorder: FC<ComponentProcessProps> = ({ id }) => {
  const { prependFileToTitle } = useTitle(id);
  const { createPath, updateFolder } = useFileSystem();
  const mediaRecorderRef = useRef<MediaRecorder | undefined>(undefined);
  const mediaStreamRef = useRef<MediaStream | undefined>(undefined);
  const chunksRef = useRef<Blob[]>([]);
  const analyserRef = useRef<AnalyserNode | undefined>(undefined);
  const animationRef = useRef<number>(0);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const [recording, setRecording] = useState(false);
  const [duration, setDuration] = useState(0);
  const [error, setError] = useState("");
  const startTimeRef = useRef(0);
  const timerRef = useRef<ReturnType<typeof setInterval> | undefined>(undefined);

  useEffect(() => () => {
      mediaStreamRef.current?.getTracks().forEach((track) => track.stop());
      mediaStreamRef.current = undefined;
      if (timerRef.current) clearInterval(timerRef.current);
      if (animationRef.current) cancelAnimationFrame(animationRef.current);
  }, []);

  const drawVisualizer = useCallback(() => {
    const canvas = canvasRef.current;
    const analyser = analyserRef.current;
    if (!canvas || !analyser) return;

    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    const bufferLength = analyser.frequencyBinCount;
    const dataArray = new Uint8Array(bufferLength);

    const draw = (): void => {
      animationRef.current = requestAnimationFrame(draw);
      analyser.getByteTimeDomainData(dataArray);

      ctx.fillStyle = "#1a1a1a";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.lineWidth = 2;
      ctx.strokeStyle = "#4fc3f7";

      ctx.beginPath();
      const sliceWidth = canvas.width / bufferLength;
      let x = 0;

      for (let i = 0; i < bufferLength; i++) {
        const v = dataArray[i] / 128;
        const y = (v * canvas.height) / 2;
        if (i === 0) {
          ctx.moveTo(x, y);
        } else {
          ctx.lineTo(x, y);
        }
        x += sliceWidth;
      }
      ctx.lineTo(canvas.width, canvas.height / 2);
      ctx.stroke();
    };

    draw();
  }, []);

  const startRecording = useCallback(async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        audio: {
          channelCount: 1,
          echoCancellation: true,
          sampleRate: 44100,
        },
      });
      mediaStreamRef.current = stream;

      const audioContext = new AudioContext();
      const source = audioContext.createMediaStreamSource(stream);
      const analyser = audioContext.createAnalyser();
      analyser.fftSize = 256;
      source.connect(analyser);
      analyserRef.current = analyser;
      drawVisualizer();

      chunksRef.current = [];
      const mimeType = MediaRecorder.isTypeSupported("audio/webm;codecs=opus")
        ? "audio/webm;codecs=opus"
        : "audio/webm";

      const recorder = new MediaRecorder(stream, { mimeType });
      mediaRecorderRef.current = recorder;

      recorder.addEventListener("dataavailable", (event) => {
        if (event.data?.size) chunksRef.current.push(event.data);
      });

      recorder.addEventListener("stop", async () => {
        if (chunksRef.current.length === 0) return;
        const blob = new Blob(chunksRef.current, { type: mimeType });
        const buffer = await blobToBuffer(blob);
        const ext = "webm";
        const name = `Recording_${Date.now()}.${ext}`;
        const filePath = await createPath(name, DESKTOP_PATH, buffer);
        await updateFolder(DESKTOP_PATH, basename(filePath));
        prependFileToTitle(basename(filePath));
        chunksRef.current = [];

        audioContext.close();
        stream.getTracks().forEach((track) => track.stop());
        mediaStreamRef.current = undefined;
        analyserRef.current = undefined;
        if (animationRef.current) cancelAnimationFrame(animationRef.current);
      });

      recorder.start();
      setRecording(true);
      setError("");
      startTimeRef.current = Date.now();
      timerRef.current = setInterval(() => {
        setDuration(Math.floor((Date.now() - startTimeRef.current) / 1000));
      }, 200);
    } catch {
      setError("Microphone access denied or not available");
    }
  }, [createPath, drawVisualizer, prependFileToTitle, updateFolder]);

  const stopRecording = useCallback(() => {
    if (
      mediaRecorderRef.current &&
      mediaRecorderRef.current.state !== "inactive"
    ) {
      mediaRecorderRef.current.stop();
    }
    setRecording(false);
    if (timerRef.current) {
      clearInterval(timerRef.current);
      timerRef.current = undefined;
    }
    setDuration(0);
  }, []);

  const formatTime = useCallback((seconds: number) => {
    const m = Math.floor(seconds / 60);
    const s = seconds % 60;
    return `${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`;
  }, []);

  const actions = useMemo(
    () => [
      {
        disabled: recording,
        label: "Record",
        onClick: startRecording,
      },
      {
        disabled: !recording,
        label: "Stop",
        onClick: stopRecording,
      },
    ],
    [recording, startRecording, stopRecording]
  );

  return (
    <StyledVoiceRecorder>
      <canvas ref={canvasRef} height={120} width={400} />
      <div className="timer">{formatTime(duration)}</div>
      {error && <div className="error">{error}</div>}
      <nav>
        {actions.map(({ disabled, label, onClick }) => (
          <Button key={label} disabled={disabled} onClick={onClick}>
            {label}
          </Button>
        ))}
      </nav>
    </StyledVoiceRecorder>
  );
};

export default memo(VoiceRecorder);
