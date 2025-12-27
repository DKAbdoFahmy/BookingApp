[app]
title = BookingTool
package.name = bookingtool
package.domain = org.jood
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,ttf,json,pkl
version = 0.1
requirements = python3,kivy==2.3.0,requests,urllib3,beautifulsoup4,openpyxl,arabic-reshaper,python-bidi,openssl
orientation = portrait
fullscreen = 0
android.permissions = INTERNET,WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE,ACCESS_NETWORK_STATE
android.api = 33
android.minapi = 21
android.accept_sdk_license = True
android.archs = arm64-v8a
p4a.branch = master

[buildozer]
log_level = 2
warn_on_root = 0
