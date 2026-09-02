import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  build: {
    outDir: "dist",
    chunkSizeWarningLimit: 1200,
    rollupOptions: {
      output: {
        // Firebase + recharts + xlsx are large and change rarely; split them into
        // their own chunks so the app chunk stays small and vendors cache well.
        manualChunks: {
          firebase: ["firebase/app", "firebase/auth", "firebase/firestore"],
          recharts: ["recharts"],
          xlsx: ["xlsx"],
        },
      },
    },
  },
});
