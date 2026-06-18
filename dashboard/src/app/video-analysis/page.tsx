"use client";

import { useState } from "react";
import {
  FileVideo,
  UploadCloud,
  Loader2,
  AlertCircle,
  CheckCircle,
  ShieldAlert,
  ShieldCheck,
  ExternalLink,
  X,
  Download,
} from "lucide-react";

interface AlertResult {
  id: string;
  message: string;
  timestamp: string;
  image_path: string;
  camera_id: string;
}

interface AnalyzeResponse {
  status: string;
  frames_processed: number;
  alert_count: number;
  alerts: AlertResult[];
  message?: string;
  output_video?: string | null;
}

export default function VideoAnalysisPage() {
  const [videoFile, setVideoFile] = useState<File | null>(null);
  const [notify, setNotify] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [message, setMessage] = useState<{ text: string; type: "success" | "error" } | null>(null);
  const [results, setResults] = useState<AnalyzeResponse | null>(null);

  const apiBaseUrl = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000";

  const formatTime = (ts: string) => {
    if (!ts || ts.length !== 15) return ts;
    const year = ts.slice(0, 4);
    const month = ts.slice(4, 6);
    const day = ts.slice(6, 8);
    const hour = ts.slice(9, 11);
    const min = ts.slice(11, 13);
    const sec = ts.slice(13, 15);
    return `${year}-${month}-${day} ${hour}:${min}:${sec}`;
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setVideoFile(e.target.files[0]);
    }
  };

  const handleRemoveVideo = () => {
    setVideoFile(null);
    setResults(null);
    setMessage(null);
    const input = document.getElementById("video-input") as HTMLInputElement | null;
    if (input) input.value = "";
  };

  const handleDownload = async () => {
    if (!results?.output_video) return;
    try {
      const res = await fetch(`${apiBaseUrl}/${results.output_video}`);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "analyzed_video.mp4";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      console.error(err);
      setMessage({ text: "Could not download the analyzed video.", type: "error" });
    }
  };

  const handleAnalyze = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!videoFile) {
      setMessage({ text: "Please select a video file to analyze.", type: "error" });
      return;
    }

    setSubmitting(true);
    setMessage(null);
    setResults(null);

    const formData = new FormData();
    formData.append("file", videoFile);
    formData.append("notify", notify ? "true" : "false");

    try {
      const res = await fetch(`${apiBaseUrl}/video/analyze`, {
        method: "POST",
        body: formData,
      });
      const data: AnalyzeResponse = await res.json();

      if (res.ok && data.status === "success") {
        setResults(data);
        setMessage({
          text:
            data.alert_count > 0
              ? `Analysis complete — ${data.alert_count} alert(s) detected.`
              : "Analysis complete — no suspicious activity detected.",
          type: "success",
        });
      } else {
        setMessage({ text: data.message || "Analysis failed.", type: "error" });
      }
    } catch (err) {
      console.error(err);
      setMessage({ text: "Connection error. Ensure the backend is running.", type: "error" });
    } finally {
      setSubmitting(false);
    }
  };

  const badgeClass = (msg: string) =>
    msg.includes("THEFT") || msg.includes("CRIMINAL")
      ? "bg-danger/20 text-danger border border-danger/20"
      : msg.includes("BLACKLIST") || msg.includes("RESTRICTED")
      ? "bg-orange-500/20 text-orange-400 border border-orange-500/20"
      : "bg-blue-500/20 text-blue-400 border border-blue-500/20";

  return (
    <div className="max-w-6xl mx-auto pb-10">
      <header className="mb-8">
        <h2 className="text-3xl font-bold tracking-tight mb-2">Video Analysis</h2>
        <p className="text-foreground/60">
          Upload a recorded clip and run it through the theft-detection model to flag suspicious activity.
        </p>
      </header>

      {message && (
        <div
          className={`mb-6 p-4 rounded-lg flex items-center gap-3 border ${
            message.type === "success"
              ? "bg-green-500/10 border-green-500/30 text-green-400"
              : "bg-danger/10 border-danger/30 text-danger"
          }`}
        >
          {message.type === "success" ? <CheckCircle className="w-5 h-5" /> : <AlertCircle className="w-5 h-5" />}
          <span className="text-sm font-medium">{message.text}</span>
        </div>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Upload Section */}
        <div className="lg:col-span-1">
          <div className="glass-panel p-6 sticky top-6">
            <div className="flex items-center gap-3 mb-6">
              <div className="p-2 rounded bg-brand/20 text-brand">
                <FileVideo className="w-5 h-5" />
              </div>
              <h3 className="text-xl font-semibold">Upload Video</h3>
            </div>

            <form onSubmit={handleAnalyze} className="space-y-4">
              <div className="space-y-2">
                <label className="text-sm font-medium text-foreground/80">Video File</label>
                <div className="mt-1 flex justify-center px-6 pt-5 pb-6 border-2 border-glass-border border-dashed rounded-lg bg-black/20 hover:bg-black/30 transition-colors cursor-pointer relative group">
                  <div className="space-y-1 text-center">
                    <UploadCloud className="mx-auto h-12 w-12 text-foreground/45 group-hover:text-brand transition-colors" />
                    <div className="flex text-sm text-foreground/60 justify-center">
                      <span className="relative rounded-md font-semibold text-brand hover:text-brand/80 focus-within:outline-none">
                        Upload a file
                      </span>
                    </div>
                    <p className="text-xs text-foreground/45">MP4, AVI, MOV up to ~200MB</p>
                  </div>
                  <input
                    id="video-input"
                    type="file"
                    accept="video/*"
                    onChange={handleFileChange}
                    className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                    required
                  />
                </div>
                {videoFile && (
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-xs text-brand font-medium break-all">Selected: {videoFile.name}</span>
                    <button
                      type="button"
                      onClick={handleRemoveVideo}
                      disabled={submitting}
                      className="flex items-center gap-1 text-xs text-foreground/50 hover:text-danger transition-colors cursor-pointer shrink-0 disabled:opacity-50"
                      title="Remove video"
                    >
                      <X className="w-3.5 h-3.5" />
                      Remove
                    </button>
                  </div>
                )}
              </div>

              <div className="flex items-start gap-2 pt-1">
                <input
                  type="checkbox"
                  id="notify-toggle"
                  checked={notify}
                  onChange={(e) => setNotify(e.target.checked)}
                  className="mt-0.5 rounded bg-black/40 border-glass-border text-brand"
                />
                <label htmlFor="notify-toggle" className="text-xs text-foreground/70 leading-relaxed">
                  Send Telegram / email alerts for detections
                  <span className="block text-foreground/40">Uses your Settings → Telegram / Email configuration.</span>
                </label>
              </div>

              <button
                type="submit"
                disabled={submitting}
                className="w-full mt-4 flex items-center justify-center gap-2 bg-brand hover:bg-brand/90 disabled:opacity-50 text-white py-2 rounded-lg font-medium transition-colors cursor-pointer"
              >
                {submitting ? <Loader2 className="w-4 h-4 animate-spin" /> : <ShieldAlert className="w-4 h-4" />}
                {submitting ? "Analyzing..." : "Analyze Video"}
              </button>

              {submitting && (
                <p className="text-xs text-foreground/50 text-center">
                  Processing frames through the AI model. Longer clips can take a while.
                </p>
              )}
            </form>
          </div>
        </div>

        {/* Results Section */}
        <div className="lg:col-span-2">
          <div className="glass-panel p-6 min-h-[300px]">
            <div className="flex items-center gap-3 mb-6">
              <div className="p-2 rounded bg-purple-500/20 text-purple-400">
                <ShieldAlert className="w-5 h-5" />
              </div>
              <h3 className="text-xl font-semibold">Detection Results</h3>
            </div>

            {results?.output_video && (
              <div className="mb-5">
                <video
                  src={`${apiBaseUrl}/${results.output_video}`}
                  controls
                  className="w-full rounded-lg bg-black border border-glass-border"
                />
                <div className="flex justify-end mt-2">
                  <button
                    onClick={handleDownload}
                    className="flex items-center gap-2 px-4 py-2 bg-brand/20 border border-brand/35 hover:bg-brand/30 text-brand rounded-lg transition-colors text-sm font-bold cursor-pointer"
                  >
                    <Download className="w-4 h-4" />
                    Download analyzed video
                  </button>
                </div>
              </div>
            )}

            {submitting ? (
              <div className="flex flex-col items-center justify-center h-[220px] text-foreground/60">
                <Loader2 className="w-8 h-8 animate-spin text-brand mb-3" />
                Analyzing video...
              </div>
            ) : !results ? (
              <div className="text-center p-12 text-foreground/40 border border-glass-border border-dashed rounded-lg bg-black/10">
                Upload a video to see detection results here.
              </div>
            ) : results.alert_count === 0 ? (
              <div className="flex flex-col items-center justify-center p-10 text-center">
                <div className="p-3 rounded-full bg-green-500/10 border border-green-500/30 text-green-400 mb-4">
                  <ShieldCheck className="w-8 h-8" />
                </div>
                <h4 className="text-lg font-semibold text-green-400 mb-1">No suspicious activity detected</h4>
                <p className="text-sm text-foreground/50">
                  Processed {results.frames_processed} frames — nothing flagged.
                </p>
              </div>
            ) : (
              <>
                <div className="mb-5 p-4 rounded-lg bg-danger/10 border border-danger/30 text-danger flex items-center gap-3">
                  <ShieldAlert className="w-5 h-5" />
                  <span className="text-sm font-medium">
                    {results.alert_count} alert(s) flagged across {results.frames_processed} frames.
                  </span>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {results.alerts.map((alert) => (
                    <div
                      key={alert.id}
                      className="glass-panel p-3 border-t-4 border-t-danger/70 hover:bg-white/[0.02] transition-all"
                    >
                      <a
                        href={`${apiBaseUrl}/${alert.image_path}`}
                        target="_blank"
                        rel="noreferrer"
                        className="block relative group"
                      >
                        {/* eslint-disable-next-line @next/next/no-img-element */}
                        <img
                          src={`${apiBaseUrl}/${alert.image_path}`}
                          alt={alert.message}
                          className="w-full h-40 object-cover rounded-md bg-black/40"
                        />
                        <span className="absolute top-2 right-2 p-1 rounded bg-black/60 text-white opacity-0 group-hover:opacity-100 transition-opacity">
                          <ExternalLink className="w-4 h-4" />
                        </span>
                      </a>
                      <div className="mt-3">
                        <span className={`inline-block px-2 py-1 rounded text-xs font-semibold ${badgeClass(alert.message)}`}>
                          {alert.message}
                        </span>
                        <p className="text-xs text-foreground/50 mt-2">{formatTime(alert.timestamp)}</p>
                      </div>
                    </div>
                  ))}
                </div>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
