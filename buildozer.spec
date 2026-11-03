[app]

title = Расписание
package.name = scheduleapp
# domain + name = org.schedule.scheduleapp — этот пакет используется в Java-коде виджета!
package.domain = org.schedule

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json
source.exclude_dirs = java,res,bin,.github,.buildozer,venv,__pycache__
source.exclude_patterns = generate_excel.py,*.xlsx,README.md

version = 1.0.0

# pandas намеренно не включён: в APK не нужен (разбор Excel делает openpyxl),
# а сборка pandas/numpy под Android очень тяжёлая и ломкая.
requirements = python3,kivy==2.3.0,openpyxl,et_xmlfile,pyjnius,android

orientation = portrait
fullscreen = 0

# --- Android ---
android.api = 33
android.minapi = 24
android.ndk = 25b
android.archs = arm64-v8a
android.accept_sdk_license = True
android.debug_artifact = apk

# python-for-android клонируется и дополняется в workflow (см. build.yml):
# в манифест добавляются receiver и activity виджета
p4a.source_dir = ~/p4a-src

# Виджет: Java-классы и ресурсы
android.add_src = java
android.add_resources = res/layout/widget_layout.xml:layout/widget_layout.xml, res/drawable/widget_bg.xml:drawable/widget_bg.xml, res/xml/schedule_widget_info.xml:xml/schedule_widget_info.xml

[buildozer]
log_level = 2
warn_on_root = 1
