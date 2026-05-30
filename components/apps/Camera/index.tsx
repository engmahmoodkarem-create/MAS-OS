import { basename } from "path";
import {
  memo,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import StyledCamera from "components/apps/Camera/StyledCamera";
import { type ComponentProcessProps } from "components/system/Apps/RenderComponent";
import useTitle from "components/system/Window/useTitle";
import { useFileSystem } from "contexts/fileSystem";
import { useProcesses } from "contexts/process";
import Button from "styles/common/Button";
import { PICTURES_FOLDER, VIDEOS_FOLDER } from "utils/constants";
import { blobToBuffer } from "utils/functions";

const Camera: FC<ComponentProcessProps> = ({ id }) => {
  const { processes: { [id]: process } = {} } = useProcesses();
  const { closing = false } = process || {};
  const { prependFileToTitle } = useTitle(id);
  const { createPath, exists, mkdir, updateFolder } =
    useFileSystem();
  const videoRef = useRef<HTMLVideoElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const mediaStreamRef = useRef<MediaStream | undefined>(undefined);
  const mediaRecorderRef = useRef<MediaRecorder | null>(null);
  const chunksRef = useRef<Blob[]>([]);
  const [recording, setRecording] = useState(false);
  const [hasStream, setHasStream] = useState(false);
  const [error, setError] = useState("");
  const [facingMode, setFacingMode] = useState<"user" | "environment">("user");

  const startStream = useCallback(async (facing: "user" | "environment") => {
    try {
      if (mediaStreamRef.current) {
        mediaStreamRef.current.getTracks().forEach((track) => track.stop());
      }
      const stream = await navigator.mediaDevices.getUserMedia({
        audio: true,
        video: { facingMode: facing, height: { ideal: 720 }, width: { ideal: 1280 } },
      });
      mediaStreamRef.current = stream;
      if (videoRef.current) {
        videoRef.current.srcObject = stream;
      }
      setHasStream(true);
      setError("");
    } catch {
      setError("Camera access denied or not available");
      setHasStream(false);
    }
  }, []);

  useEffect(() => {
    if (!closing) {
      startStream(facingMode);
    }
    return () => {
      mediaStreamRef.current?.getTracks().forEach((track) => track.stop());
      mediaStreamRef.current = undefined;
    };
  }, [closing, facingMode, startStream]);

  const switchCamera = useCallback(() => {
    setFacingMode((prev) => (prev === "user" ? "environment" : "user"));
  }, []);

  const capturePhoto = useCallback(() => {
    const video = videoRef.current;
    const canvas = canvasRef.current;
    if (!video || !canvas || !hasStream) return;

    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    const ctx = canvas.getContext("2d");
    if (!ctx) return;

    ctx.drawImage(video, 0, 0);
    canvas.toBlob((blob) => {
      if (blob) {
        blobToBuffer(blob).then(async (buffer) => {
          const ext = "png";
          const name = `Photo_${Date.now()}.${ext}`;
          if (!(await exists(PICTURES_FOLDER))) {
            await mkdir(PICTURES_FOLDER);
          }
          const filePath = await createPath(name, PICTURES_FOLDER, buffer);
          await updateFolder(PICTURES_FOLDER, basename(filePath));
          prependFileToTitle(basename(filePath));
        });
      }
    }, "image/png");
  }, [createPath, exists, hasStream, mkdir, prependFileToTitle, updateFolder]);

  const startRecording = useCallback(() => {
    const stream = mediaStreamRef.current;
    if (!stream) return;

    chunksRef.current = [];
    const mimeType = MediaRecorder.isTypeSupported("video/webm;codecs=vp9,opus")
      ? "video/webm;codecs=vp9,opus"
      : MediaRecorder.isTypeSupported("video/webm;codecs=vp8,opus")
        ? "video/webm;codecs=vp8,opus"
        : "video/webm";

    const recorder = new MediaRecorder(stream, { mimeType });
    mediaRecorderRef.current = recorder;

    recorder.addEventListener("dataavailable", (event) => {
      if (event.data?.size) {
        chunksRef.current.push(event.data);
      }
    });

    recorder.addEventListener("stop", async () => {
      if (chunksRef.current.length === 0) return;
      const blob = new Blob(chunksRef.current, { type: mimeType });
      const buffer = await blobToBuffer(blob);
      const ext = "webm";
      const name = `Video_${Date.now()}.${ext}`;
      if (!(await exists(VIDEOS_FOLDER))) {
        await mkdir(VIDEOS_FOLDER);
      }
      const filePath = await createPath(name, VIDEOS_FOLDER, buffer);
      await updateFolder(VIDEOS_FOLDER, basename(filePath));
      prependFileToTitle(basename(filePath));
      chunksRef.current = [];
    });

    recorder.start();
    setRecording(true);
  }, [createPath, exists, mkdir, prependFileToTitle, updateFolder]);

  const stopRecording = useCallback(() => {
    if (
      mediaRecorderRef.current &&
      mediaRecorderRef.current.state !== "inactive"
    ) {
      mediaRecorderRef.current.stop();
    }
    setRecording(false);
  }, []);

  const actions = useMemo(
    () => [
      {
        disabled: !hasStream,
        label: "Capture Photo",
        onClick: capturePhoto,
      },
      {
        disabled: !hasStream || recording,
        label: "Start Recording",
        onClick: startRecording,
      },
      {
        disabled: !recording,
        label: "Stop Recording",
        onClick: stopRecording,
      },
      {
        disabled: !hasStream,
        label: `Switch to ${facingMode === "user" ? "Rear" : "Front"} Camera`,
        onClick: switchCamera,
      },
    ],
    [capturePhoto, facingMode, hasStream, recording, startRecording, stopRecording, switchCamera]
  );

  return (
    <StyledCamera>
      <canvas ref={canvasRef} />
      {error ? (
        <div className="error">{error}</div>
      ) : (
        <video ref={videoRef} autoPlay muted playsInline />
      )}
      <nav>
        {actions.map(({ disabled, label, onClick }) => (
          <Button key={label} disabled={disabled} onClick={onClick}>
            {label}
          </Button>
        ))}
      </nav>
    </StyledCamera>
  );
};

export default memo(Camera);
