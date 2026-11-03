package org.schedule.scheduleapp;

import android.app.PendingIntent;
import android.appwidget.AppWidgetManager;
import android.appwidget.AppWidgetProvider;
import android.content.Context;
import android.content.Intent;
import android.content.SharedPreferences;
import android.text.SpannableStringBuilder;
import android.text.Spanned;
import android.text.style.ForegroundColorSpan;
import android.text.style.RelativeSizeSpan;
import android.text.style.StyleSpan;
import android.graphics.Typeface;
import android.widget.RemoteViews;

import org.json.JSONArray;
import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

/**
 * Виджет расписания. Данные читаются из files/widget_data.json, который
 * пишет Python-приложение. Настройки каждого виджета (группа, подгруппа,
 * день) лежат в SharedPreferences "widgets" с ключами w<ID>_*.
 */
public class ScheduleWidgetProvider extends AppWidgetProvider {

    static final String ACT_PREV = "org.schedule.scheduleapp.WIDGET_PREV";
    static final String ACT_NEXT = "org.schedule.scheduleapp.WIDGET_NEXT";
    static final String[] DAYS = {"Пн", "Вт", "Ср", "Чт", "Пт", "Сб"};

    @Override
    public void onUpdate(Context c, AppWidgetManager m, int[] ids) {
        for (int id : ids) updateWidget(c, m, id);
    }

    @Override
    public void onReceive(Context c, Intent intent) {
        String a = intent.getAction();
        if (ACT_PREV.equals(a) || ACT_NEXT.equals(a)) {
            int id = intent.getIntExtra(AppWidgetManager.EXTRA_APPWIDGET_ID,
                    AppWidgetManager.INVALID_APPWIDGET_ID);
            if (id != AppWidgetManager.INVALID_APPWIDGET_ID) {
                SharedPreferences p = prefs(c);
                int shift = p.getInt("w" + id + "_shift", 0) + (ACT_NEXT.equals(a) ? 1 : -1);
                p.edit().putInt("w" + id + "_shift", shift).apply();
                updateWidget(c, AppWidgetManager.getInstance(c), id);
            }
        } else {
            super.onReceive(c, intent);
        }
    }

    @Override
    public void onDeleted(Context c, int[] ids) {
        SharedPreferences.Editor e = prefs(c).edit();
        for (int id : ids) {
            e.remove("w" + id + "_grp").remove("w" + id + "_sub")
                    .remove("w" + id + "_day").remove("w" + id + "_shift");
        }
        e.apply();
    }

    static SharedPreferences prefs(Context c) {
        return c.getSharedPreferences("widgets", Context.MODE_PRIVATE);
    }

    static JSONObject loadData(Context c) {
        try {
            File f = new File(c.getFilesDir(), "widget_data.json");
            if (!f.exists()) return null;
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try (FileInputStream in = new FileInputStream(f)) {
                byte[] buf = new byte[16384];
                int n;
                while ((n = in.read(buf)) > 0) bos.write(buf, 0, n);
            }
            return new JSONObject(bos.toString("UTF-8"));
        } catch (Exception e) {
            return null;
        }
    }

    /** Индекс текущего дня: Пн=0 … Вс=6. */
    static int todayIndex() {
        return (Calendar.getInstance().get(Calendar.DAY_OF_WEEK) + 5) % 7;
    }

    /** Номер недели относительно anchor (понедельник недели «над чертой»). */
    static int weekIndex(String anchor) {
        try {
            String[] a = anchor.split("-");
            Calendar ac = Calendar.getInstance();
            ac.clear();
            ac.set(Integer.parseInt(a[0]), Integer.parseInt(a[1]) - 1, Integer.parseInt(a[2]));
            Calendar nc = Calendar.getInstance();
            nc.add(Calendar.DAY_OF_MONTH, -todayIndex());
            nc.set(Calendar.HOUR_OF_DAY, 0);
            nc.set(Calendar.MINUTE, 0);
            nc.set(Calendar.SECOND, 0);
            nc.set(Calendar.MILLISECOND, 0);
            long days = Math.round((nc.getTimeInMillis() - ac.getTimeInMillis()) / 86400000.0);
            return (int) Math.floor(days / 7.0);
        } catch (Exception e) {
            return 0;
        }
    }

    static int idOf(Context c, String name) {
        return c.getResources().getIdentifier(name, "id", c.getPackageName());
    }

    static PendingIntent actionIntent(Context c, String action, int id, int req) {
        Intent i = new Intent(c, ScheduleWidgetProvider.class);
        i.setAction(action);
        i.putExtra(AppWidgetManager.EXTRA_APPWIDGET_ID, id);
        return PendingIntent.getBroadcast(c, id * 4 + req, i,
                PendingIntent.FLAG_UPDATE_CURRENT | PendingIntent.FLAG_IMMUTABLE);
    }

    static void append(SpannableStringBuilder sb, String text, int color, boolean bold,
                       float size) {
        int s = sb.length();
        sb.append(text);
        int e = sb.length();
        sb.setSpan(new ForegroundColorSpan(color), s, e, Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);
        if (bold) sb.setSpan(new StyleSpan(Typeface.BOLD), s, e, Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);
        if (size != 1f) sb.setSpan(new RelativeSizeSpan(size), s, e, Spanned.SPAN_EXCLUSIVE_EXCLUSIVE);
    }

    static void updateWidget(Context c, AppWidgetManager m, int id) {
        String pkg = c.getPackageName();
        int layout = c.getResources().getIdentifier("widget_layout", "layout", pkg);
        RemoteViews v = new RemoteViews(pkg, layout);
        SharedPreferences p = prefs(c);
        JSONObject data = loadData(c);

        // клики: ‹ › и открытие приложения по заголовку
        v.setOnClickPendingIntent(idOf(c, "widget_prev"), actionIntent(c, ACT_PREV, id, 1));
        v.setOnClickPendingIntent(idOf(c, "widget_next"), actionIntent(c, ACT_NEXT, id, 2));
        Intent launch = c.getPackageManager().getLaunchIntentForPackage(pkg);
        if (launch != null) {
            v.setOnClickPendingIntent(idOf(c, "widget_title"), PendingIntent.getActivity(
                    c, id * 4 + 3, launch,
                    PendingIntent.FLAG_UPDATE_CURRENT | PendingIntent.FLAG_IMMUTABLE));
        }

        if (data == null || data.optJSONArray("lessons") == null
                || data.optJSONArray("lessons").length() == 0) {
            v.setTextViewText(idOf(c, "widget_title"), "Расписание");
            v.setTextViewText(idOf(c, "widget_body"),
                    "Нет данных.\nОткройте приложение и импортируйте расписание.");
            m.updateAppWidget(id, v);
            return;
        }

        // кто / что показывать
        String faculty = "", group = "";
        int course = 0, sub = 0;
        JSONObject d = data.optJSONObject("defaults");
        if (d != null) {
            faculty = d.optString("faculty");
            course = d.optInt("course");
            group = d.optString("group");
            sub = d.optInt("subgroup");
        }
        String grp = p.getString("w" + id + "_grp", null);
        if (grp != null) {
            String[] parts = grp.split("\\|", -1);
            if (parts.length == 3) {
                faculty = parts[0];
                try { course = Integer.parseInt(parts[1]); } catch (Exception ignored) { }
                group = parts[2];
                sub = p.getInt("w" + id + "_sub", 0);
            }
        }
        int fixedDay = p.getInt("w" + id + "_day", -1);
        int shift = p.getInt("w" + id + "_shift", 0);
        int today = todayIndex();
        boolean auto = fixedDay < 0;
        int base = auto ? (today == 6 ? 0 : today) : fixedDay;
        int day = (((base + shift) % 6) + 6) % 6;
        int wk = weekIndex(data.optString("anchor", "2000-01-03"))
                + ((auto && today == 6) ? 1 : 0);
        boolean over = (((wk % 2) + 2) % 2) == 0;
        String weekKey = over ? "over" : "under";

        // выбрать занятия
        List<JSONObject> list = new ArrayList<>();
        JSONArray all = data.optJSONArray("lessons");
        for (int i = 0; i < all.length(); i++) {
            JSONObject l = all.optJSONObject(i);
            if (l == null) continue;
            if (!faculty.equals(l.optString("f")) || course != l.optInt("c")
                    || !group.equals(l.optString("g")) || day != l.optInt("d")) continue;
            String w = l.optString("w");
            if (!w.equals("both") && !w.equals(weekKey)) continue;
            int sg = l.optInt("sg");
            if (sub != 0 && sg != 0 && sg != sub) continue;
            list.add(l);
        }
        Collections.sort(list, new Comparator<JSONObject>() {
            @Override
            public int compare(JSONObject a, JSONObject b) {
                int r = Integer.compare(a.optInt("p"), b.optInt("p"));
                return r != 0 ? r : Integer.compare(a.optInt("sg"), b.optInt("sg"));
            }
        });

        String title = group + " · " + DAYS[day] + " · " + (over ? "над чертой" : "под чертой")
                + (sub != 0 ? " · " + sub + " подгр." : "");
        v.setTextViewText(idOf(c, "widget_title"), title);

        SpannableStringBuilder sb = new SpannableStringBuilder();
        if (list.isEmpty()) {
            append(sb, "Занятий нет", 0xFF8B949E, false, 1.1f);
        }
        for (JSONObject l : list) {
            if (sb.length() > 0) sb.append("\n");
            int pair = l.optInt("p");
            String head = (pair > 0 ? pair + " · " : "") + l.optString("s") + "–" + l.optString("e");
            append(sb, head + "\n", 0xFF60A5FA, true, 0.9f);
            String subj = l.optString("sj");
            String type = l.optString("t");
            append(sb, subj + (type.isEmpty() ? "" : " (" + type + ")") + "\n",
                    0xFFE6EDF3, true, 1f);
            String te = l.optString("te");
            String r = l.optString("r");
            String det = te + ((!te.isEmpty() && !r.isEmpty()) ? " · " : "")
                    + (r.isEmpty() ? "" : "ауд. " + r);
            if (!det.isEmpty()) append(sb, det + "\n", 0xFF8B949E, false, 0.85f);
        }
        v.setTextViewText(idOf(c, "widget_body"), sb);
        m.updateAppWidget(id, v);
    }
}
