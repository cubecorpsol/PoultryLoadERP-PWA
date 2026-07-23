# PoultryLoad ERP — Android APK build project

This folder wraps the PoultryLoad ERP web app (already built, in `www/`) in a
[Capacitor](https://capacitorjs.com) shell so it can be compiled into a real,
installable Android `.apk`.

Why this last step can't be done for you automatically: producing an `.apk`
requires the Android SDK, Android Gradle Plugin, and Gradle itself, which are
multi-gigabyte downloads from Google's and Gradle's servers. That has to
happen on a machine with Android Studio installed — it can't be run inside
this chat.

## What you need on your machine

- [Node.js](https://nodejs.org) 18+
- [Android Studio](https://developer.android.com/studio) (includes the Android SDK)

## Steps

1. Unzip this project, open a terminal in the folder, and install dependencies:
   ```
   npm install
   ```

2. Add the native Android project (generates an `android/` folder):
   ```
   npm run android:add
   ```

3. Open it in Android Studio:
   ```
   npm run android:open
   ```

4. In Android Studio: **Build → Build Bundle(s) / APK(s) → Build APK(s)**.
   The signed/debug `.apk` will appear under
   `android/app/build/outputs/apk/debug/app-debug.apk` — copy that file to
   your phone (or use **Run ▶** to install it on a connected device/emulator
   directly).

5. To make future changes: edit `www/` (or rebuild it from the React source
   with `esbuild`), then run `npx cap sync android` and rebuild in Android
   Studio.

## Notes

- The app stores its data locally on the device (localStorage), so nothing
  is sent to a server — this remains a front-end-only prototype. The "MongoDB
  + Node.js + Express" backend from the original brief is not included; the
  app works standalone.
- For a production release (Play Store), you'll need to generate a signing
  key and build a **release** APK/AAB instead of the debug one — Android
  Studio's **Build → Generate Signed Bundle / APK** wizard handles this.
- If you'd rather skip Android Studio entirely, see the alternative
  "no-code" path in the chat message this project was shared in
  (PWABuilder.com).
