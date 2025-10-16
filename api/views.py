import pandas as pd
import io
from rest_framework.views import APIView
from rest_framework.parsers import MultiPartParser
from django.http import HttpResponse
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import math

class ProcessExcelView(APIView):
    parser_classes = [MultiPartParser]

    def post(self, request):
        f = request.FILES.get('file')
        output_type = request.POST.get('output_type', 'excel')

        if not f:
            return HttpResponse('لم يتم رفع الملف', status=400)

        try:
            df = pd.read_excel(f)
        except Exception as e:
            return HttpResponse(f'خطأ في قراءة الملف: {e}', status=400)

        rename_map = {
            'Monthly Income (DZD)': 'monthly_income',
            'Debt Ratio (%)': 'debt_ratio',
            'Loan Duration (months)': 'loan_duration_months',
            'Annual Interest Rate (%)': 'annual_interest_rate_percent'
        }
        df = df.rename(columns=rename_map)

        required = list(rename_map.values())
        missing = [c for c in required if c not in df.columns]
        if missing:
            return HttpResponse(f'الأعمدة الناقصة: {missing}', status=400)

        df = df.fillna(0)

        results = []
        for _, row in df.iterrows():
            m_income = row['monthly_income']
            ratio = row['debt_ratio'] / 100 if row['debt_ratio'] > 1 else row['debt_ratio']
            months = row['loan_duration_months']
            rate = row['annual_interest_rate_percent']

            if any(math.isnan(v) for v in [m_income, ratio, months, rate]):
                continue

            # ======= Calculs principaux =======
            monthly_rate = rate / 100 / 12
            mensualite_max = m_income * ratio
            if monthly_rate == 0:
                loan_amount = mensualite_max * months
            else:
                loan_amount = mensualite_max * ((1 + monthly_rate) * months - 1) / (monthly_rate * (1 + monthly_rate) * months)

            # arrondir proprement
            mensualite_max = round(mensualite_max, 2)
            loan_amount = round(loan_amount, 0)

            # ======= Résultats =======
            results.append({
                'Taux d’intérêt (%)': f"{rate:.2f}",
                'Durée (année)': round(months / 12, 2),
                'Durée en mois': int(months),
                'Différé': 1,
                'Revenu mensuel (DZD)': int(m_income),
                'Taux d’endettement (%)': f"{ratio * 100:.0f}",
                'Mensualité Maximale (DZD)': mensualite_max,
                'Montant Crédit (DZD)': loan_amount
            })

        if not results:
            return HttpResponse('Aucune ligne valide trouvée dans le fichier.', status=400)

        df_out = pd.DataFrame(results)

        # ======= PDF =======
        if output_type == 'pdf':
            pdf_buffer = io.BytesIO()
            doc = SimpleDocTemplate(
                pdf_buffer,
                pagesize=A4,
                rightMargin=40,
                leftMargin=40,
                topMargin=40,
                bottomMargin=30
            )
            styles = getSampleStyleSheet()
            elements = []

            for i, row in df_out.iterrows():
                elements.append(Paragraph("🏦 Détails du crédit", styles['Title']))
                elements.append(Spacer(1, 12))

                data = [[k, str(v)] for k, v in row.items()]
                table = Table(data, colWidths=[180, 180])
                table.setStyle(TableStyle([
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('BACKGROUND', (0, 0), (-1, -1), colors.whitesmoke),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 10),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
                ]))

                elements.append(table)
                elements.append(Spacer(1, 20))

                if i < len(df_out) - 1:
                    elements.append(PageBreak())

            doc.build(elements)
            pdf_buffer.seek(0)
            response = HttpResponse(pdf_buffer.read(), content_type='application/pdf')
            response['Content-Disposition'] = 'attachment; filename="loan_report.pdf"'
            return response

        # ======= Excel =======
        else:
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_out.to_excel(writer, index=False)
            excel_buffer.seek(0)
            response = HttpResponse(
                excel_buffer.read(),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            response['Content-Disposition'] = 'attachment; filename="loan_results.xlsx"'
            return response
