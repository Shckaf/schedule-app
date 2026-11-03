package org.schedule.scheduleapp;

import android.app.Activity;
import android.appwidget.AppWidgetManager;
import android.content.Context;
import android.content.Intent;
import android.os.Bundle;
import android.view.View;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.LinearLayout;
import android.widget.Spinner;
import android.widget.TextView;

import org.json.JSONArray;
import org.json.JSONObject;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

/** Экран настройки виджета: группа, подгруппа и день недели. */
public class WidgetConfigActivity extends Activity {

    private int widgetId = AppWidgetManager.INVALID_APPWIDGET_ID;

    private int dp(int v) {
        return (int) (v * getResources().getDisplayMetrics().density);
    }

    private TextView label(String text) {
        TextView t = new TextView(this);
        t.setText(text);
        t.setTextSize(13);
        t.setPadding(0, dp(12), 0, dp(2));
        return t;
    }

    @Override
    protected void onCreate(Bundle b) {
        super.onCreate(b);
        setResult(RESULT_CANCELED);
        Bundle extras = getIntent().getExtras();
        if (extras != null) {
            widgetId = extras.getInt(AppWidgetManager.EXTRA_APPWIDGET_ID,
                    AppWidgetManager.INVALID_APPWIDGET_ID);
        }
        if (widgetId == AppWidgetManager.INVALID_APPWIDGET_ID) {
            finish();
            return;
        }

        // список групп из данных приложения
        final List<String> keys = new ArrayList<>();
        final List<String> labels = new ArrayList<>();
        JSONObject data = ScheduleWidgetProvider.loadData(this);
        if (data != null && data.optJSONArray("lessons") != null) {
            TreeMap<String, String> uniq = new TreeMap<>();
            JSONArray ls = data.optJSONArray("lessons");
            for (int i = 0; i < ls.length(); i++) {
                JSONObject l = ls.optJSONObject(i);
                if (l == null) continue;
                String key = l.optString("f") + "|" + l.optInt("c") + "|" + l.optString("g");
                if (!uniq.containsKey(key)) {
                    uniq.put(key, l.optString("g") + " · " + l.optInt("c") + " курс · "
                            + l.optString("f"));
                }
            }
            for (Map.Entry<String, String> e : uniq.entrySet()) {
                keys.add(e.getKey());
                labels.add(e.getValue());
            }
        }

        LinearLayout root = new LinearLayout(this);
        root.setOrientation(LinearLayout.VERTICAL);
        root.setPadding(dp(20), dp(16), dp(20), dp(16));

        TextView title = new TextView(this);
        title.setText("Настройка виджета");
        title.setTextSize(20);
        root.addView(title);

        if (keys.isEmpty()) {
            TextView msg = new TextView(this);
            msg.setText("Расписание ещё не загружено. Откройте приложение, "
                    + "импортируйте Excel и добавьте виджет снова.");
            msg.setPadding(0, dp(16), 0, dp(16));
            root.addView(msg);
            Button close = new Button(this);
            close.setText("Закрыть");
            close.setOnClickListener(new View.OnClickListener() {
                @Override
                public void onClick(View v) { finish(); }
            });
            root.addView(close);
            setContentView(root);
            return;
        }

        root.addView(label("Группа"));
        final Spinner grpSp = new Spinner(this);
        grpSp.setAdapter(new ArrayAdapter<>(this,
                android.R.layout.simple_spinner_dropdown_item, labels));
        root.addView(grpSp);

        root.addView(label("Подгруппа"));
        final Spinner subSp = new Spinner(this);
        subSp.setAdapter(new ArrayAdapter<>(this, android.R.layout.simple_spinner_dropdown_item,
                new String[]{"Все подгруппы", "1 подгруппа", "2 подгруппа"}));
        root.addView(subSp);

        root.addView(label("День недели"));
        final Spinner daySp = new Spinner(this);
        daySp.setAdapter(new ArrayAdapter<>(this, android.R.layout.simple_spinner_dropdown_item,
                new String[]{"Автоматически (сегодня)", "Понедельник", "Вторник", "Среда",
                        "Четверг", "Пятница", "Суббота"}));
        root.addView(daySp);

        // предвыбор: группа, выбранная в приложении
        JSONObject d = data.optJSONObject("defaults");
        if (d != null) {
            String cur = d.optString("faculty") + "|" + d.optInt("course") + "|"
                    + d.optString("group");
            int pos = keys.indexOf(cur);
            if (pos >= 0) grpSp.setSelection(pos);
            subSp.setSelection(Math.max(0, Math.min(2, d.optInt("subgroup"))));
        }

        Button ok = new Button(this);
        ok.setText("Добавить виджет");
        ok.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Context c = WidgetConfigActivity.this;
                c.getSharedPreferences("widgets", MODE_PRIVATE).edit()
                        .putString("w" + widgetId + "_grp", keys.get(grpSp.getSelectedItemPosition()))
                        .putInt("w" + widgetId + "_sub", subSp.getSelectedItemPosition())
                        .putInt("w" + widgetId + "_day", daySp.getSelectedItemPosition() - 1)
                        .putInt("w" + widgetId + "_shift", 0)
                        .apply();
                ScheduleWidgetProvider.updateWidget(c, AppWidgetManager.getInstance(c), widgetId);
                Intent result = new Intent();
                result.putExtra(AppWidgetManager.EXTRA_APPWIDGET_ID, widgetId);
                setResult(RESULT_OK, result);
                finish();
            }
        });
        LinearLayout.LayoutParams lp = new LinearLayout.LayoutParams(
                LinearLayout.LayoutParams.MATCH_PARENT, LinearLayout.LayoutParams.WRAP_CONTENT);
        lp.topMargin = dp(20);
        root.addView(ok, lp);
        setContentView(root);
    }
}
