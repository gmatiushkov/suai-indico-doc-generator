import psycopg2
from datetime import datetime
import json
from collections import defaultdict

def format_date_rus(dt):
    """Форматирование даты на русский."""
    months = {
        'January': 'января', 'February': 'февраля', 'March': 'марта', 'April': 'апреля',
        'May': 'мая', 'June': 'июня', 'July': 'июля', 'August': 'августа',
        'September': 'сентября', 'October': 'октября', 'November': 'ноября', 'December': 'декабря'
    }
    day = dt.strftime("%d").lstrip("0")
    month = months[dt.strftime("%B")]
    return f"{day} {month}"

def create_conference_json(json_file_path):
    """Создание JSON файла с данными всех конференций из базы данных."""
    # Подключение к базе данных Indico
    conn = psycopg2.connect(
        dbname="indico",
        user="user",
        password="user",
        host="localhost",
        port="5432"
    )
    cur = conn.cursor()

    # Словарь для маппинга состояний рецензирования
    review_states = {
        0: "not submitted",
        1: "submitted",
        2: "accepted",
        3: "rejected",
        4: "to be corrected"
    }

    # Структура для хранения данных
    conference_data = {
        "conferences": []
    }

    # Извлечение всех конференций
    cur.execute("""
        SELECT
            id,
            title,
            start_dt AT TIME ZONE 'UTC' AT TIME ZONE 'Europe/Moscow' AS start_dt_moscow,
            end_dt AT TIME ZONE 'UTC' AT TIME ZONE 'Europe/Moscow' AS end_dt_moscow,
            venue_name,
            room_name,
            address,
            timezone
        FROM events.events
        WHERE is_deleted = false
        ORDER BY id;
    """)
    events = cur.fetchall()

    for event in events:
        event_id, title, start_dt, end_dt, venue_name, room_name, address, timezone = event
        conference = {
            "id": event_id,
            "title": title,
            "start_date": format_date_rus(start_dt),
            "start_time": start_dt.strftime("%H:%M"),
            "end_date": format_date_rus(end_dt),
            "end_time": end_dt.strftime("%H:%M"),
            "venue_name": venue_name or "",
            "room_name": room_name or "",
            "address": address or "",
            "timezone": timezone,
            "sessions": []
        }

        # --- Извлечение оргкомитета (leadership) ---
        cur.execute("""
            SELECT
                r.name AS role_name,
                u.first_name,
                u.last_name,
                u.affiliation,
                ue.email
            FROM events.roles r
            JOIN events.role_members rm ON rm.role_id = r.id
            JOIN users.users u ON rm.user_id = u.id
            LEFT JOIN users.emails ue ON ue.user_id = u.id
            WHERE r.event_id = %s
            ORDER BY r.id;
        """, (event_id,))
        roles = cur.fetchall()

        leadership = {}
        for role_name, first_name, last_name, affiliation, email in roles:
            person_data = {
                "name": f"{last_name} {first_name}",
                "affiliation": affiliation or ""
            }
            if email:
                person_data["email"] = email

            if "Научный руководитель" in role_name:
                leadership["scientific_leader"] = person_data
            elif "Зам" in role_name:
                leadership["deputy_leader"] = person_data
            elif "Секретарь" in role_name:
                leadership["secretary"] = person_data
            else:
                leadership[role_name] = person_data

        conference["leadership"] = leadership

        # Извлечение данных о сессиях и докладах
        cur.execute("""
            SELECT
                sb.id,
                s.title AS session_title,
                t.start_dt AT TIME ZONE 'UTC' AT TIME ZONE 'Europe/Moscow' AS start_dt_moscow,
                sb.duration,
                e.room_name
            FROM events.session_blocks sb
            JOIN events.sessions s ON sb.session_id = s.id
            JOIN events.timetable_entries t ON t.session_block_id = sb.id
            JOIN events.events e ON s.event_id = e.id
            WHERE s.event_id = %s AND t.event_id = %s AND t.type = 1
            ORDER BY t.start_dt;
        """, (event_id, event_id))
        sessions = cur.fetchall()

        for session_index, session in enumerate(sessions, 1):
            session_id, session_title, start_dt, duration, room_name = session
            session_number = str(session_index)
            
            session_data = {
                "id": session_id,
                "number": session_number,
                "title": session_title,
                "date": format_date_rus(start_dt),
                "start_time": start_dt.strftime("%H:%M"),
                "duration": str(duration),
                "room_name": f"{room_name} БМ." if room_name else "",
                "contributions": []
            }
            
            cur.execute("""
                SELECT
                    c.id,
                    c.title AS contribution_title,
                    t.start_dt AT TIME ZONE 'UTC' AT TIME ZONE 'Europe/Moscow' AS start_dt_moscow,
                    c.duration,
                    p.first_name,
                    p.last_name,
                    p.affiliation,
                    COALESCE(r.state, 0) AS review_state
                FROM events.timetable_entries t
                JOIN events.contributions c ON t.contribution_id = c.id
                LEFT JOIN events.sessions s ON c.session_id = s.id
                JOIN events.contribution_person_links cp ON cp.contribution_id = c.id
                JOIN events.persons p ON cp.person_id = p.id
                LEFT JOIN event_paper_reviewing.revisions r ON r.contribution_id = c.id
                WHERE t.event_id = %s AND t.type = 2 AND cp.is_speaker = true
                AND s.title = %s
                ORDER BY t.start_dt;
            """, (event_id, session_title))
            contributions = cur.fetchall()
            
            for contrib in contributions:
                contrib_id, title, start_dt, duration, first_name, last_name, affiliation, review_state = contrib
                session_data["contributions"].append({
                    "id": contrib_id,
                    "title": title,
                    "start_time": start_dt.strftime("%H:%M"),
                    "duration": str(duration),
                    "speaker": {
                        "first_name": first_name,
                        "last_name": last_name,
                        "full_name": f"{last_name} {first_name}",
                        "affiliation": affiliation or ""
                    },
                    "review_state": review_states.get(review_state, "unknown")
                })
            
            conference["sessions"].append(session_data)

        conference_data["conferences"].append(conference)

    # Закрытие соединения
    cur.close()
    conn.close()

    # Сохранение данных в JSON
    with open(json_file_path, "w", encoding="utf-8") as f:
        json.dump(conference_data, f, ensure_ascii=False, indent=4)

