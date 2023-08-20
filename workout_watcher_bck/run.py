from setup import app, db
from flask import request, jsonify
import flask_excel as excel
from models import User, MonthlyReport, TotalCalories

import os
import traceback
import pandas as pd
from io import BytesIO
from pandas import ExcelWriter
from datetime import datetime, timedelta


# ----

def custom_excel(res, filename):
    resp = excel.Response(res)
    resp.headers["Content-Disposition"] = "attachment; filename= %s" % filename
    resp.mimetype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return resp


# ------


@app.route("/", methods=['GET'])
def ping():
    return jsonify({"message": "Yes you are connected!!"}), 200


@app.route('/user', methods=['POST'])
def add_user():
    try:
        request_body = request.get_json()
        data = request_body['data']
        user = User(username=data['username'], email=data['email'], password=data['password'])
        db.session.add(user)
        db.session.commit()
        user = User.query.filter_by(email=data['email']).first()
        if not user:
            raise Exception("Not able to add the user")
        return jsonify({"valid": True}), 200
    except Exception as error:
        print(traceback.format_exc())
        print(f"Addition of the user failed - {error}")
        return jsonify({"error": f"Addition of the user failed - {error}"}), 500


@app.route("/user", methods=['PUT'])
def validate_user():
    try:
        request_body = request.get_json()
        data = request_body['data']
        user = User.query.filter_by(email=data['email']).first()
        if user and user.check_password(data['password']):
            return jsonify({'id': user.id, 'fullname': user.username, 'email': user.email, 'valid': True}), 200
        return jsonify({'valid': False}), 404
    except Exception as error:
        print(f"Error in validating the user - {error}")
        return jsonify({"error": f"Error in validating the user - {error}"}), 500


@app.route('/submit', methods=['POST'])
def add_today_workout():
    calories_switch = {
        "Light exercises": 5,
        "Moderate exercises": 8,
        "Heavy exercises": 10
    }

    email = request.args.get('email')
    if not email:
        return jsonify({'error': 'Email is required.'}), 404

    request_body = request.get_json()
    data = request_body['data']

    # Get the current date and time
    current_date = datetime.now().date()
    try:
        # Filter the TotalCalories table to get records for the current day
        total_calories_obj = db.session.query(TotalCalories).filter(TotalCalories.email == email,
                                                                    TotalCalories.timestamp == current_date).first()

        if int(data['hours']) > 0:
            hourly_minutes = int(data['hours']) * 60
        else:
            hourly_minutes = 0

        calories = (int(data['minutes']) + hourly_minutes) * calories_switch[data['category']]

        if total_calories_obj is None:
            new_total_calories_obj = TotalCalories(email=email, total_calories=calories, timestamp=current_date)
            db.session.add(new_total_calories_obj)
        else:
            total_calories_obj.total_calories += calories

        # Create a new MonthlyReport object with the desired values
        report = MonthlyReport(exercise_category=data['category'], hours=data['hours'], minutes=data['minutes'],
                               seconds=data['seconds'], timestamp=current_date, email=email)

        # Add the new report to the database session and commit the changes
        db.session.add(report)
        db.session.commit()
    except Exception as error:
        print(f"Error while adding the data to database - {error}")
        print(traceback.format_exc())
        return jsonify({"error": f"Error while adding the data to database - {error}"}), 500
    else:
        return jsonify({"valid": True}), 200


@app.route('/monthly-reports', methods=['GET'])
def get_monthly_reports():
    email = request.args.get('email')
    if not email:
        return jsonify({'error': 'Email is required.'}), 404

    if os.path.exists('monthly_reports.xlsx'):
        os.remove('monthly_reports.xlsx')

    # Query the database for records within the date range
    # Get the monthly report from today for the given email
    today = datetime.utcnow().date()
    last_month = today - timedelta(days=30)

    try:
        monthly_report = MonthlyReport.query.filter(
            MonthlyReport.email == email,
            MonthlyReport.timestamp >= last_month,
            MonthlyReport.timestamp <= today
        ).all()

        # Convert the data to a pandas dataframe
        data = [(report.exercise_category, report.hours, report.minutes, report.seconds, report.timestamp) for report in
                monthly_report]
        df = pd.DataFrame(data, columns=['Exercise Category', 'Hours', 'Minutes', 'Seconds', 'Timestamp'])

        output = BytesIO()
        with ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Monthly Report', index=False)

            workbook = writer.book
            worksheet = writer.sheets['Monthly Report']
            money_fmt = workbook.add_format({'text_wrap': True, 'align': 'left'})
            worksheet.set_column('A:F', 30, money_fmt)

        excel_result = output.getvalue()

    except Exception as error:
        print(f"Error while fetching the monthly report - {error}")
        print(traceback.format_exc())
        return jsonify({"error": f"Error while fetching the monthly report - {error}"}), 500
    else:
        return custom_excel(excel_result, 'Monthly Report.xlsx')


@app.route('/weekly_report', methods=['GET'])
def get_weekly_total_calories():
    email = request.args.get('email')
    if not email:
        return jsonify({'error': 'Email is required.'}), 404

    # Get the current date and time
    current_time = datetime.utcnow()

    # Calculate the timestamp for 7 days ago
    past_time = current_time - timedelta(days=7)

    try:
        # Filter the TotalCalories table to get records for the past 7 days
        total_calories = TotalCalories.query.filter(TotalCalories.email == email,
                                                    TotalCalories.timestamp >= past_time).order_by(
            TotalCalories.timestamp.asc()).all()

        result = []
        times = []
        values = []
        for item in total_calories:
            got_date = item.timestamp.strftime('%Y-%m-%d').split(" ")[0]
            result.append([got_date, item.total_calories])
            times.append(got_date)
            values.append(item.total_calories)
    except Exception as error:
        print(f"Error while fetching the weekly data for analytics - {error}")
        print(traceback.format_exc())
        return jsonify({"error": f"Error while fetching the weekly data for analytics - {error}"}), 500
    else:
        return jsonify({"total": result, "times": times, "values": values}), 200


if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(port=5001)
