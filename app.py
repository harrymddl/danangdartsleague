import os
import re
import pandas as pd
from flask import Flask, render_template_string, request, redirect, flash

app = Flask(__name__)
app.secret_key = "ddl_secret_key_2026"
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

MATCH_LOG_FILE = 'match_history.csv'

# Initialize an empty history file if it doesn't exist
if not os.path.exists(MATCH_LOG_FILE):
    pd.DataFrame(columns=[
        'Player', 'Opponent', 'LegsWon', 'LegsLost', 'DartsThrown', 
        'MatchAvg', 'CoHits', 'CoAtt', 'H100', 'H140', 'H180'
    ]).to_csv(MATCH_LOG_FILE, index=False)

def extract_checkout_stats(val):
    """Parses text fields like '37.50% (3/8)' into explicit hits and attempts."""
    if pd.isna(val) or str(val).strip() == '-':
        return 0, 0
    match = re.search(r'\((\d+)/(\d+)\)', str(val))
    if match:
        return int(match.group(1)), int(match.group(2))
    return 0, 0

def parse_and_log_match(filepath):
    df = pd.read_excel(filepath)
    # Identify players dynamically from column headings
    p1, p2 = df.columns[1], df.columns[2]
    
    # Restructure into string-indexed key-value pairs
    stats = {str(row['Stats']).strip(): (row[p1], row[p2]) for _, row in df.iterrows()}
    
    # Extract structural variables safely
    l_won_1, l_won_2 = int(stats['Legs won'][0]), int(stats['Legs won'][1])
    darts_1, darts_2 = int(stats['Darts thrown'][0]), int(stats['Darts thrown'][1])
    avg_1, avg_2 = float(stats['Average'][0]), float(stats['Average'][1])
    
    co_hit_1, co_att_1 = extract_checkout_stats(stats['Checkout %'][0])
    co_hit_2, co_att_2 = extract_checkout_stats(stats['Checkout %'][1])
    
    h100_1, h100_2 = int(stats['100+'][0]), int(stats['100+'][1])
    h140_1, h140_2 = int(stats['140+'][0]), int(stats['140+'][1])
    h180_1, h180_2 = int(stats['180'][0]), int(stats['180'][1])
    
    # Build logs for both sides
    match_rows = [
        [p1, p2, l_won_1, l_won_2, darts_1, avg_1, co_hit_1, co_att_1, h100_1, h140_1, h180_1],
        [p2, p1, l_won_2, l_won_1, darts_2, avg_2, co_hit_2, co_att_2, h100_2, h140_2, h180_2]
    ]
    
    df_log = pd.DataFrame(match_rows, columns=[
        'Player', 'Opponent', 'LegsWon', 'LegsLost', 'DartsThrown', 
        'MatchAvg', 'CoHits', 'CoAtt', 'H100', 'H140', 'H180'
    ])
    df_log.to_csv(MATCH_LOG_FILE, mode='a', header=False, index=False)

def generate_league_table():
    df_history = pd.read_csv(MATCH_LOG_FILE)
    if df_history.empty:
        return pd.DataFrame()
        
    summary = []
    # Dynamic processing for all distinct players found in history
    for player in df_history['Player'].unique():
        pdf = df_history[df_history['Player'] == player]
        
        mp = len(pdf)
        w = sum(pdf['LegsWon'] > pdf['LegsLost'])
        l = mp - w
        lf = pdf['LegsWon'].sum()
        la = pdf['LegsLost'].sum()
        ld = lf - la
        pts = w  # 1 point per match win
        
        # Weighted Running Average using Total Points Scored / (Total Darts / 3)
        total_darts = pdf['DartsThrown'].sum()
        total_points_scored = sum(pdf['MatchAvg'] * (pdf['DartsThrown'] / 3))
        running_avg = (total_points_scored / (total_darts / 3)) if total_darts > 0 else 0
        
        # Weighted Checkout %
        total_co_hits = pdf['CoHits'].sum()
        total_co_att = pdf['CoAtt'].sum()
        running_co = (total_co_hits / total_co_att * 100) if total_co_att > 0 else 0
        
        summary.append({
            'Player': player.upper(), 'MP': mp, 'W': w, 'L': l, 'LF': lf, 'LA': la, 'LD': ld, 'Pts': pts,
            'Avg': round(running_avg, 2), 'CO': f"{round(running_co, 1)}%",
            '180': pdf['H180'].sum(), '140+': pdf['H140'].sum(), '100+': pdf['H100'].sum()
        })
        
    df_table = pd.DataFrame(summary)
    # Sort hierarchy: Points DESC -> Leg Difference DESC -> Legs For DESC
    df_table = df_table.sort_values(by=['Pts', 'LD', 'LF'], ascending=[False, False, False]).reset_index(drop=True)
    df_table.insert(0, 'Pos', range(1, len(df_table) + 1))
    return df_table

# --- UI Template (Combined Dashboard and Admin Upload Screen) ---
HTML_TEMPLATE = """
<!xltype html>
<html>
<head>
    <title>Danang Darts League (DDL)</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <style>
        body { background-color: #0f1115; color: #e2e8f0; font-family: system-ui, sans-serif; }
        .card { background-color: #1a1f26; border: 1px solid #2d3748; color: #fff; }
        .table { color: #e2e8f0; border-color: #2d3748; }
        th { background-color: #242b35 !important; color: #a0aec0 !important; font-weight: 600; }
        td { background-color: #1a1f26 !important; vertical-align: middle; }
        .pos-col { font-weight: bold; color: #edd045; }
        .ld-pos { color: #48bb78; } .ld-neg { color: #f56565; }
    </style>
</head>
<body class="py-5">
    <div class="container">
        <h2 class="text-center mb-4 text-uppercase tracking-wider">Danang Darts League (DDL) - Standings</h2>
        
        {% with messages = get_flashed_messages() %}
          {% if messages %}<div class="alert alert-success bg-success text-white border-0">{{ messages[0] }}</div>{% endif %}
        {% endwith %}

        <div class="card p-4 mb-5">
            <h5 class="mb-3 text-muted">Upload Weekly Match Result (.xlsx)</h5>
            <form method="POST" action="/upload" enctype="multipart/form-data" class="row g-3 align-items-center">
                <div class="col-sm-8"><input class="form-control" type="file" name="file" accept=".xlsx" required></div>
                <div class="col-sm-4"><button type="submit" class="btn btn-warning w-100 fw-bold text-uppercase">Submit Match</button></div>
            </form>
        </div>

        <div class="card p-4">
            <div class="table-responsive">
                <table class="table table-hover text-center mb-0">
                    <thead>
                        <tr>
                            <th>Pos</th><th>Player</th><th>MP</th><th>W</th><th>L</th><th>LF</th><th>LA</th><th>LD</th><th>Pts</th><th>Avg</th><th>C/O %</th><th>180</th><th>140+</th><th>100+</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for row in table %}
                        <tr>
                            <td class="pos-col">{{ row.Pos }}</td>
                            <td class="fw-bold text-start ps-4">{{ row.Player }}</td>
                            <td>{{ row.MP }}</td><td>{{ row.W }}</td><td>{{ row.L }}</td><td>{{ row.LF }}</td><td>{{ row.LA }}</td>
                            <td class="{{ 'ld-pos' if row.LD >= 0 else 'ld-neg' }} fw-bold">{{ '+' if row.LD > 0 }}{{ row.LD }}</td>
                            <td class="fw-bold text-warning">{{ row.Pts }}</td><td>{{ row.Avg }}</td><td>{{ row.CO }}</td>
                            <td>{{ row['180'] }}</td><td>{{ row['140+'] }}</td><td>{{ row['100+'] }}</td>
                        </tr>
                        {% else %}
                        <tr><td colspan="14" class="text-muted py-4">No match data logged yet. Upload your first game above!</td></tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</body>
</html>
"""

@app.route('/')
def index():
    df_table = generate_league_table()
    table_data = df_table.to_dict(orient='records') if not df_table.empty else []
    return render_template_string(HTML_TEMPLATE, table=table_data)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files: return redirect('/')
    file = request.files['file']
    if file.filename == '': return redirect('/')
    
    if file and file.filename.endswith('.xlsx'):
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        parse_and_log_match(filepath)
        flash(f"Successfully processed and merged match values from '{file.filename}'!")
        
    return redirect('/')

if __name__ == '__main__':
    app.run(debug=True, port=5000)
