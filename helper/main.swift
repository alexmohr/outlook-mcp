// Runs an ASWebAuthenticationSession (system auth session -> Enterprise SSO / Company Portal,
// conditional access works) and prints the callback URL to stdout.
// usage: auth-helper <authorize-url> <callback-scheme>
import AppKit
import AuthenticationServices

final class Anchor: NSObject, ASWebAuthenticationPresentationContextProviding {
    let window = NSWindow(contentRect: NSRect(x: 200, y: 200, width: 1, height: 1),
                          styleMask: [.borderless], backing: .buffered, defer: false)
    func presentationAnchor(for session: ASWebAuthenticationSession) -> ASPresentationAnchor { window }
}

let args = CommandLine.arguments
guard args.count >= 3, let url = URL(string: args[1]) else {
    FileHandle.standardError.write(Data("usage: auth-helper <url> <scheme>\n".utf8))
    exit(2)
}

let app = NSApplication.shared
app.setActivationPolicy(.accessory)
let anchor = Anchor()

let session = ASWebAuthenticationSession(url: url, callback: .customScheme(args[2])) { callbackURL, error in
    if let callbackURL = callbackURL {
        print(callbackURL.absoluteString)
        exit(0)
    }
    FileHandle.standardError.write(Data("\(error?.localizedDescription ?? "Auth session cancelled")\n".utf8))
    exit(1)
}
session.presentationContextProvider = anchor
session.prefersEphemeralWebBrowserSession = false

anchor.window.makeKeyAndOrderFront(nil)
app.activate(ignoringOtherApps: true)
if !session.start() {
    FileHandle.standardError.write(Data("Failed to start ASWebAuthenticationSession\n".utf8))
    exit(1)
}
DispatchQueue.main.asyncAfter(deadline: .now() + 300) {
    FileHandle.standardError.write(Data("Authentication timed out\n".utf8))
    exit(1)
}
app.run()
