import SwiftUI
import WebKit

struct AppWebView: UIViewRepresentable {
    func makeUIView(context: Context) -> WKWebView {
        let webView = WKWebView(frame: .zero)
        webView.scrollView.bounces = false

        if let localIndex = Bundle.main.url(forResource: "index", withExtension: "html", subdirectory: "WebMirror") {
            webView.loadFileURL(localIndex, allowingReadAccessTo: localIndex.deletingLastPathComponent())
        } else if let remoteURL = URL(string: "https://followamuse.nl/apps/carnaval-der-dieren/") {
            webView.load(URLRequest(url: remoteURL))
        }

        return webView
    }

    func updateUIView(_ uiView: WKWebView, context: Context) {
    }
}
