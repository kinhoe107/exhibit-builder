import { Component, StrictMode, type ErrorInfo, type ReactNode } from "react";
import { createRoot } from "react-dom/client";
import "../app/globals.css";

const root = document.getElementById("root");

if (!root) {
  throw new Error("Exhibit Builder could not find its application root.");
}
const applicationRoot = root;

type ErrorBoundaryState = { error: Error | null };

class ApplicationErrorBoundary extends Component<
  { children: ReactNode },
  ErrorBoundaryState
> {
  state: ErrorBoundaryState = { error: null };

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { error };
  }

  componentDidCatch(error: Error, info: ErrorInfo) {
    console.error("Exhibit Builder renderer error", error, info.componentStack);
  }

  render() {
    if (this.state.error) {
      return (
        <main className="startup-failure" role="alert">
          <p className="eyebrow">EXHIBIT BUILDER</p>
          <h1>The workspace could not be displayed.</h1>
          <p>
            No document content has been changed. Close and reopen Exhibit
            Builder, then retry the analysis.
          </p>
          <button type="button" onClick={() => window.location.reload()}>
            Reload workspace
          </button>
        </main>
      );
    }
    return this.props.children;
  }
}

function showStartupFailure(caught: unknown) {
  const detail = caught instanceof Error ? caught.message : String(caught);
  console.error("Exhibit Builder could not start", caught);
  applicationRoot.innerHTML = "";
  const failure = document.createElement("main");
  failure.className = "startup-failure";
  failure.setAttribute("role", "alert");
  failure.innerHTML = `
    <p class="eyebrow">EXHIBIT BUILDER</p>
    <h1>The workspace could not be opened.</h1>
    <p>No document content has been changed. Close and reopen Exhibit Builder, then retry.</p>
    <button type="button">Reload workspace</button>
  `;
  failure.querySelector("button")?.addEventListener("click", () => window.location.reload());
  applicationRoot.append(failure);
  document.title = "Exhibit Builder | Startup error";
  void detail;
}

async function mountApplication() {
  // Let the inline boot screen paint before evaluating the full workspace
  // bundle. This keeps first launch responsive on slower managed Windows
  // machines without changing the offline document-processing path.
  await new Promise<void>((resolve) => {
    // Hidden automated smoke windows may not receive animation frames, so use
    // a timer as the portable first-turn yield for both visible and hidden
    // renderers.
    window.setTimeout(resolve, 0);
  });
  const { default: BundleBuilder } = await import("../app/BundleBuilder");
  createRoot(applicationRoot).render(
    <StrictMode>
      <ApplicationErrorBoundary>
        <BundleBuilder />
      </ApplicationErrorBoundary>
    </StrictMode>,
  );
}

void mountApplication().catch(showStartupFailure);
