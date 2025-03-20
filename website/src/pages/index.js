import useDocusaurusContext from "@docusaurus/useDocusaurusContext";
import Layout from "@theme/Layout";
import React, { useState } from "react";
import Link from '@docusaurus/Link';
import styles from "./styles.module.css";

function VideoPlayer() {
  const [showVideo, setShowVideo] = useState(false);
  
  const playVideo = () => {
    setShowVideo(true);
  };
  
  return (
    <div className={styles.videoContainer}>
      {!showVideo ? (
        <div className={styles.thumbnailContainer} onClick={playVideo}>
          <img 
            className={styles.thumbnail} 
            src="https://img.youtube.com/vi/7v52C9WZaxY/maxresdefault.jpg" 
            alt="Lokka Demo Video" 
          />
          <div className={styles.playButtonContainer}>
            <div className={styles.playButtonOuter}>
              <div className={styles.playButtonInner}>
                <div className={styles.playIcon}></div>
              </div>
            </div>
          </div>
        </div>
      ) : (
        <iframe 
          className={styles.videoFrame}
          src="https://www.youtube.com/embed/7v52C9WZaxY?autoplay=1"
          title="Lokka Demo"
          frameBorder="0"
          allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture"
          allowFullScreen
        ></iframe>
      )}
    </div>
  );
}

export default function Home() {
  const { siteConfig } = useDocusaurusContext();
  return (
    <Layout
      title="Lokka"
      description="Beyond Commands, Beyond Clicks. A glimpse into the future of Managing Microsoft 365!">
      <main>
        <div className={styles.hero}>
          <div className={styles.container}>
            <div className={styles.heroContent}>
              <h1 className={styles.heroTitle}>Lokka</h1>
              <p className={styles.heroSubtitle}>Lokka is an AI agent tool that brings the power of Microsoft Graph to AI agents like GitHub Copilot and Claude that run on your local desktop.</p>
                <p className={styles.heroSubtitle}>Get a glimpse into the future of administering Microsoft 365 👇</p>
            </div>
            <VideoPlayer />
            <div className={styles.buttonContainer}>
              <Link
                className={styles.tryButton}
                to="/docs/intro">
                Try Lokka
              </Link>
            </div>
          </div>
        </div>
      </main>
    </Layout>
  );
}
