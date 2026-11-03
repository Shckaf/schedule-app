package org.schedule.scheduleapp;

import android.appwidget.AppWidgetManager;
import android.content.ComponentName;
import android.content.Context;
import android.net.Uri;
import android.os.Build;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

/** Статические методы, которые вызывает Python-код (через pyjnius). */
public class ScheduleBridge {

    /** Перерисовать все размещённые виджеты расписания. */
    public static void refreshAll(Context c) {
        AppWidgetManager m = AppWidgetManager.getInstance(c);
        int[] ids = m.getAppWidgetIds(new ComponentName(c, ScheduleWidgetProvider.class));
        for (int id : ids) {
            ScheduleWidgetProvider.updateWidget(c, m, id);
        }
    }

    /** Попросить лаунчер добавить виджет на рабочий стол (Android 8+). */
    public static boolean requestPin(Context c) {
        if (Build.VERSION.SDK_INT < 26) return false;
        AppWidgetManager m = AppWidgetManager.getInstance(c);
        if (!m.isRequestPinAppWidgetSupported()) return false;
        return m.requestPinAppWidget(
                new ComponentName(c, ScheduleWidgetProvider.class), null, null);
    }

    /** Скопировать файл по content:// Uri в обычный файл приложения. */
    public static boolean copyUri(Context c, Uri uri, String dest) {
        try (InputStream in = c.getContentResolver().openInputStream(uri);
             OutputStream out = new FileOutputStream(dest)) {
            byte[] buf = new byte[16384];
            int n;
            while ((n = in.read(buf)) > 0) {
                out.write(buf, 0, n);
            }
            return true;
        } catch (Exception e) {
            return false;
        }
    }
}
