import datetime
from datetime import datetime as dt
from decimal import Decimal
import json
from django.views.generic import View, ListView, TemplateView
from django.contrib.auth.decorators import login_required
from django.utils.decorators import method_decorator
from django.http import JsonResponse, Http404, HttpResponse
from financial.utils import Render
from django.utils import timezone
from django.views.decorators.csrf import csrf_exempt
from django.http import HttpResponse
from companyparameter.models import Companyparameter
from accountspayable.models import Apmain, Apdetail
from checkvoucher.models import Cvmain
from journalvoucher.models import Jvmain
from officialreceipt.models import Ormain
from subledger.models import Subledger
from customer.models import Customer
from chartofaccount.models import Chartofaccount
from collections import namedtuple
from django.db import connection
from django.template.loader import render_to_string
import pandas as pd
import io
import xlsxwriter


@method_decorator(login_required, name='dispatch')
class IndexView(TemplateView):
    template_name = 'nontrade/index.html'
    
    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('nontrade.view_nontrade'):
            raise Http404
        return super(TemplateView, self).dispatch(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        context = super(TemplateView, self).get_context_data(**kwargs)
        context['chart'] = Chartofaccount.objects.filter(isdeleted=0, accounttype='P', nontrade='Y').order_by('accountcode')

        return context


@csrf_exempt
def tagarnontrade(request):
    ar_nontrade = []
    try:
        ar_nontrade = json.loads(request.POST.get('ar_nontrade'))
        
        main = ar_nontrade[0]
        main_id = main['sl_id']
        main_document_type = main['documentType']
        main_document_number = main['documentNum']
        main_balance_code = main['balanceCode']
        main_amount = main['amount']
        
        total = ar_nontrade[2]['total']
        computed_balance = Decimal(ar_nontrade[3]['computed_balance'])
        
        breakdown = ar_nontrade[1]['breakdown']
        msg = 'Invalid'
        status = 0
        if main_document_type == 'AP':

            main = Apmain.objects.filter(isdeleted=0, apnum=main_document_number).first()
            if main:
                print 'Valid AP Transaction'
                tdate = main.apdate
                status = 1
            else:
                msg = 'Invalid AP Transaction'
                status = 0
        elif main_document_type == 'CV':

            main = Cvmain.objects.filter(isdeleted=0, cvnum=main_document_number).first()
            if main:
                print 'Valid CV Transaction'
                tdate = main.cvdate
                status = 1
            else:
                msg = 'Invalid CV Transaction'
                status = 0
        elif main_document_type == 'JV':
            
            main = Jvmain.objects.filter(isdeleted=0, jvnum=main_document_number).first()
            if main:
                print 'Valid JV Transaction'
                tdate = main.jvdate
                status = 1
            else:
                msg = 'Invalid JV Transaction'
                status = 0
        elif main_document_type == 'OR':
            
            main = Ormain.objects.filter(isdeleted=0, ornum=main_document_number).first()
            if main:
                print 'Valid OR Transaction'
                tdate = main.ordate
                status = 1
            else:
                msg = 'Invalid OR Transaction'
                status = 0

        if status == 0:
            response = {
                'status': 'failed',
                'message': msg
            }
        else:
            if main_balance_code == 'Credit':

                main_exp = Subledger.objects.filter(isdeleted=0, id=main_id).first()
                main_exp.document_reftype = main_document_type
                main_exp.document_refnum = main_exp.document_num
                main_exp.document_refamount = computed_balance
                main_exp.document_refdate = tdate
                main_exp.tag_id = main_exp.pk
                main_exp.is_closed = 1 if float(computed_balance) == 0.0 else 0
                print 'computed_balance', computed_balance, main_exp.is_closed
                main_exp.save()
                
                for exp in breakdown:
                    sub = Subledger.objects.filter(isdeleted=0, id=exp['sl_id']).first()

                    sub.document_reftype = main_document_type
                    sub.document_refnum = main_document_number
                    sub.document_refdate = tdate
                    sub.tag_id = main_exp.pk
                    sub.is_closed = 1
                    sub.save()

                response = {
                    'status': 'success'
                }
            elif main_balance_code == 'Debit':
                
                main_exp = Subledger.objects.filter(isdeleted=0, id=main_id).first()
                main_exp.document_reftype = main_document_type
                main_exp.document_refnum = main_exp.document_num
                main_exp.document_refamount = computed_balance
                main_exp.document_refdate = tdate
                main_exp.tag_id = main_exp.pk
                main_exp.is_closed = 1 if float(computed_balance) == 0.0 else 0
                print 'computed_balance debit', computed_balance, main_exp.is_closed
                main_exp.save()
                
                for exp in breakdown:
                    sub = Subledger.objects.filter(isdeleted=0, id=exp['sl_id']).first()

                    sub.document_reftype = main_document_type
                    sub.document_refnum = main_document_number
                    sub.document_refdate = tdate
                    sub.tag_id = main_exp.pk
                    sub.is_closed = 1
                    sub.save()

                response = {
                    'status': 'success'
                }
            else:
                response = {
                    'status': 'failed',
                    'message': 'Invalid data!'
                }
    
    except Exception as e:
        print 'error', e
        response = {
            'status': 'failed',
            'message': str(e)
        }

    return JsonResponse(response)
    
    
# @csrf_exempt
# def tagging(request):

#     # print request.GET["id"]
#     # print request.GET["reftype"]
#     # print request.GET["refnum"]
#     # print request.GET["refdate"]

#     msg = 'Successfully tag'
#     status = 1
#     tdate= ''

#     if request.GET["reftype"] == '':
#         msg = 'Reftype is empty'
#         status = 0
#     elif request.GET["refnum"] == '':
#         msg = 'Refnum is empty'
#         status = 0
#     # elif request.GET["refdate"] == '':
#     #     msg = 'Refdate is empty'
#     #     status = 0
#     else:
#         sub = Subledger.objects.filter(isdeleted=0, id=request.GET["id"]).first()
#         if request.GET["reftype"] == 'AP':
#             print 'AP'
#             main = Apmain.objects.filter(isdeleted=0, apnum=request.GET["refnum"]).first()
#             if main:
#                 msg = 'Valid AP Transaction'
#                 tdate = main.apdate
#                 status = 1
#             else:
#                 msg = 'Invalid AP Transaction'
#                 status = 0
#         elif request.GET["reftype"] == 'CV':
#             print 'CV'
#             main = Cvmain.objects.filter(isdeleted=0, cvnum=request.GET["refnum"]).first()
#             if main:
#                 msg = 'Valid CV Transaction'
#                 tdate = main.cvdate
#                 status = 1
#             else:
#                 msg = 'Invalid CV Transaction'
#                 status = 0
#         elif request.GET["reftype"] == 'JV':
#             print 'JV'
#             main = Jvmain.objects.filter(isdeleted=0, jvnum=request.GET["refnum"]).first()
#             if main:
#                 msg = 'Valid JV Transaction'
#                 tdate = main.jvdate
#                 status = 1
#             else:
#                 msg = 'Invalid JV Transaction'
#                 status = 0
#         elif request.GET["reftype"] == 'OR':
#             print 'OR'
#             main = Ormain.objects.filter(isdeleted=0, ornum=request.GET["refnum"]).first()
#             if main:
#                 msg = 'Valid OR Transaction'
#                 tdate = main.ordate
#                 status = 1
#             else:
#                 msg = 'Invalid OR Transaction'
#                 status = 0


#         if status == 1:
#             # print 'Valid Transaction'
#             # if sub.document_type != request.GET["reftype"]:
#             #     msg = 'Invalid Transaction Tagging Subledger Type: '+ sub.document_type + ' vs Ref Type: '+ request.GET["reftype"]
#             #     status = 0
#             msg = 'Successfully tag'
#             sub.document_reftype = request.GET["reftype"]
#             sub.document_refnum = request.GET["refnum"]
#             sub.document_refdate = tdate
#             sub.save()
#             status = 1

#             print sub.document_type
#             print status

#     data = {
#         'status': status,
#         'msg': msg,
#         'tdate': tdate
#     }


#     return JsonResponse(data)

#@csrf_exempt
#@method_decorator(login_required, name='dispatch')
@login_required
def transgenerate(request):
    
    transactions = {
        '1': 'A/R Non-Trade',
        '2': 'A/P Non-Trade'
    }
    report_types = {
        '1': 'Subsidiary Ledger',
        '2': 'Schedule',
        '3': 'Statement of Account'
    }

    dto = request.GET["dto"]
    dfrom = request.GET["dfrom"]
    transaction = request.GET["transaction"]
    chartofaccount = request.GET["chartofaccount"]
    payeecode = request.GET["payeecode"]
    payeename = request.GET["payeename"]
    report = request.GET["report"]

    viewhtml = ''
    context = {}
    if report == '1':
        print "subsidiary ledger"

        data = queryLedger(dto, dfrom, transaction, chartofaccount, payeecode, payeename, isNT=False, isClosedOnly=False)

        tdebit = 0
        tcredit = 0
        for val in data:
            tdebit += val.debitamount
            tcredit += val.creditamount

        context['data'] = data
        context['tdebit'] = tdebit
        context['tcredit'] = tcredit
        context['transaction'] = transactions[transaction]
        context['reporttype'] = report_types[report]
        viewhtml = render_to_string('nontrade/transaction_result_ledger.html', context)
    elif report == '2':
        print "schedule"

        data = queryScheduled(dto, dfrom, transaction, chartofaccount, payeecode, payeename)

        tbalx = 0
        tbal = 0
        tbalbeg = 0
        tbaltran = 0
        for val in data:
            tbal += val.balamount
            tbalx += val.balamountx
            tbalbeg += val.begamt
            tbaltran += val.extrabal

        context['data'] = data
        context['tbalx'] = tbalx
        context['tbal'] = tbal
        context['tbalbeg'] = tbalbeg
        context['tbaltran'] = tbaltran
        context['dfrom'] = dfrom
        context['transaction'] = transactions[transaction]
        context['reporttype'] = report_types[report]
        viewhtml = render_to_string('nontrade/transaction_result_scheduled.html', context)

    elif report == '3':
        print "statement of account"

        data = querySOA(dto, dfrom, transaction, chartofaccount, payeecode, payeename)
        transdata = querySOATran(dfrom, '2019-01-01', transaction, chartofaccount, payeecode, payeename)

        chart = Chartofaccount.objects.filter(isdeleted=0, accounttype='P', nontrade='Y', id=chartofaccount).first()

        print chart.balancecode

        beg_code = 'D'
        beg = 0
        end = 0

        # Begging Balance
        if transaction == '1':
            cust = Customer.objects.filter(isdeleted=0, code=payeecode).first()
            if cust:
                begdata = begBalance(cust.id)

                if begdata:
                    for i in begdata:
                        #print(i.id)
                        beg = i.beg_amt
                        beg_code = i.beg_code
            # print type(begdata)
            # print str(begdata['beg_amt'])
            # print begdata.beg_code
            # print begdata.beg_date

        if beg_code != chart.balancecode:
            beg = beg * -1

        datalist = {}
        counter = 0
        amount = 0
        balance = beg
        transdebit = 0
        transcredit = 0
        dft = pd.DataFrame(transdata)
        for index, row in dft.iterrows():
            if row['balancecode'] == 'D':
                amount = (row['debitamount'])
                transdebit += (row['debitamount'])
            else:
                amount = (row['creditamount'])
                transcredit += (row['creditamount'])

            if row['balancecode'] != chart.balancecode:
                amount = amount * -1

            balance = balance + amount
            end = balance


        transbeg = transdebit - transcredit
        df = pd.DataFrame(data)
        for index, row in df.iterrows():

            if row['balancecode'] == 'D':
                amount = (row['debitamount'])
            else:
                amount = (row['creditamount'])

            if row['balancecode'] != chart.balancecode:
                amount = amount * -1

            balance = balance + amount
            datalist[counter] = dict(document_date=row['document_date'], document_type=row['document_type'], document_num=row['document_num'],
                                        particulars=row['particulars'],
                                        debitamount=float(format(row['debitamount'], '.2f')),
                                        creditamount=float(format(row['creditamount'], '.2f')),
                                        balance=float(format(balance, '.2f')))
            end = balance
            counter += 1
        #print datalist
        context['datalist'] = datalist
        context['beg'] = beg
        context['dfrom'] = dfrom
        context['transbeg'] = transbeg
        context['end'] = end
        context['transaction'] = transactions[transaction]
        context['reporttype'] = report_types[report]
        viewhtml = render_to_string('nontrade/transaction_result_soa.html', context)

    data = {
        'status': 'success',
        'viewhtml': viewhtml,
    }
    return JsonResponse(data)

@method_decorator(login_required, name='dispatch')
class GeneratePDF(View):
    def get(self, request):
        company = Companyparameter.objects.all().first()
        dto = request.GET["dto"]
        dfrom = request.GET["dfrom"]
        transaction = request.GET["transaction"]
        chartofaccount = request.GET["chartofaccount"]
        payeecode = request.GET["payeecode"]
        payeename = request.GET["payeename"]
        report = request.GET["report"]
        tbal = 0
        begbal = 0
        chart = Chartofaccount.objects.filter(isdeleted=0, accounttype='P', nontrade='Y', id=chartofaccount).first()

        if report == '2':
            title = 'SCHEDULE OF ACCOUNTS PAYABLE (NON-TRADE)'

            data = queryScheduled(dto, dfrom, transaction, chartofaccount, payeecode, payeename)

            for val in data:
                tbal += val.balamount

            context = {
                "title": title,
                "today": timezone.now(),
                "company": company,
                "listing": data,
                "tbal": tbal,
                "chartofaccount": chart,
                "begbal": begbal,
                "datefrom": datetime.datetime.strptime(dfrom, '%Y-%m-%d'),
                "dateto": datetime.datetime.strptime(dto, '%Y-%m-%d'),
                "username": request.user,
            }

            return Render.render('nontrade/pdf_scheduled.html', context)
        elif report == '3':
            print "statement of account"

            title = 'STATEMENT OF ACCOUNT'

            data = querySOA(dto, dfrom, transaction, chartofaccount, payeecode, payeename)
            transdata = querySOATran(dfrom, '2019-01-01', transaction, chartofaccount, payeecode, payeename)


            chart = Chartofaccount.objects.filter(isdeleted=0, accounttype='P', nontrade='Y', id=chartofaccount).first()

            print chart.balancecode

            beg_code = 'D'
            beg = 1000
            end = 0

            # Begging Balance
            if transaction == '1':
                cust = Customer.objects.filter(isdeleted=0, code=payeecode).first()
                if cust:
                    begdata = begBalance(cust.id)

                    if begdata:
                        for i in begdata:
                            # print(i.id)
                            beg = i.beg_amt
                            beg_code = i.beg_code



            if beg_code != chart.balancecode:
                beg = beg * -1

            datalist = {}
            counter = 0
            amount = 0
            balance = beg
            transdebit = 0
            transcredit = 0
            cusup = ""
            dft = pd.DataFrame(transdata)
            for index, row in dft.iterrows():
                if row['balancecode'] == 'D':
                    amount = (row['debitamount'])
                    transdebit += (row['debitamount'])
                else:
                    amount = (row['creditamount'])
                    transcredit += (row['creditamount'])

                if row['balancecode'] != chart.balancecode:
                    amount = amount * -1

                balance = balance + amount
                end = balance

            transbeg = transdebit - transcredit
            df = pd.DataFrame(data)

            for index, row in df.iterrows():

                if row['balancecode'] == 'D':
                    amount = (row['debitamount'])
                else:
                    amount = (row['creditamount'])

                if row['balancecode'] != chart.balancecode:
                    amount = amount * -1

                balance = balance + amount
                cusup = row['pcode']+' - '+row['pname']
                datalist[counter] = dict(document_date=row['document_date'], document_type=row['document_type'],
                                        document_num=row['document_num'],
                                        particulars=row['particulars'],
                                        debitamount=float(format(row['debitamount'], '.2f')),
                                        creditamount=float(format(row['creditamount'], '.2f')),
                                        balance=float(format(balance, '.2f')))
                end = balance
                counter += 1
            context = {
                "title": title,
                "today": timezone.now(),
                "company": company,
                "listing": datalist,
                "balance": balance,
                "chartofaccount": chart,
                "beg": beg,
                "end": end,
                "cusup": cusup,
                "transbeg": transbeg,
                "datefrom": datetime.datetime.strptime(dfrom, '%Y-%m-%d'),
                "dateto": datetime.datetime.strptime(dto, '%Y-%m-%d'),
                "username": request.user,
            }
            return Render.render('nontrade/pdf_soa.html', context)



@method_decorator(login_required, name='dispatch')
class TransExcel(View):
    def get(self, request):

        dto = request.GET["dto"]
        dfrom = request.GET["dfrom"]
        transaction = request.GET["transaction"]
        chartofaccount = request.GET["chartofaccount"]
        payeecode = request.GET["payeecode"]
        payeename = request.GET["payeename"]
        report = request.GET["report"]


        output = io.BytesIO()

        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()

        # variables
        bold = workbook.add_format({'bold': 1})
        formatdate = workbook.add_format({'num_format': 'MM/DD/YYYY'})
        centertext = workbook.add_format({'bold': 1, 'align': 'center'})

        if report == '1':
            print "subsidiary ledger"
            filename = "subsidiary_ledger.xlsx"

            result = queryLedger(dto, dfrom, transaction, chartofaccount, payeecode, payeename, isNT=False, isClosedOnly=False)

            tdebit = 0
            tcredit = 0

            # title
            worksheet.write('A1', 'NON-TRADE SUBSIDIARY LEDGER', bold)
            worksheet.write('A2', 'AS OF ' + str(dfrom) + ' to ' + str(dto), bold)

            # header
            worksheet.write('A4', 'Date', bold)
            worksheet.write('B4', 'Type', bold)
            worksheet.write('C4', 'Number', bold)
            worksheet.write('D4', 'Customer / Supplier', bold)
            worksheet.write('E4', 'Remarks', bold)
            worksheet.write('F4', 'Debit Amount', bold)
            worksheet.write('G4', 'Credit Amount', bold)
            worksheet.write('H4', 'Ref Type', bold)
            worksheet.write('I4', 'Ref No', bold)
            worksheet.write('J4', 'Ref Date', bold)

            row = 5
            col = 0

            # print result

            for data in result:
                worksheet.write(row, col, data.document_date, formatdate)
                worksheet.write(row, col + 1, data.document_type)
                if data.orsource == 'A':
                    worksheet.write(row, col + 2, str('OR') + '' + data.document_num)
                elif data.orsource == 'C':
                    worksheet.write(row, col + 2, str('CR') + '' + data.document_num)
                else:
                    worksheet.write(row, col + 2, data.document_num)
                if data.pcode:
                    worksheet.write(row, col + 3, data.pcode+'-'+data.pname)
                else:
                    worksheet.write(row, col + 3, 'N/A - NO CUSTOMER/SUPPLIER')
                worksheet.write(row, col + 4, data.particulars)
                worksheet.write(row, col + 5, float(format(data.debitamount, '.2f')))
                worksheet.write(row, col + 6, float(format(data.creditamount, '.2f')))
                worksheet.write(row, col + 7, data.document_reftype)
                worksheet.write(row, col + 8, data.document_refnum)
                worksheet.write(row, col + 9, data.document_refdate, formatdate)

                tdebit += data.debitamount
                tcredit += data.creditamount

                row += 1

            worksheet.write(row, col + 4, 'Total')
            worksheet.write(row, col + 5, float(format(tdebit, '.2f')))
            worksheet.write(row, col + 6, float(format(tcredit, '.2f')))

        elif report == '2':
            filename = "schedule.xlsx"

            result = queryScheduled(dto, dfrom, transaction, chartofaccount, payeecode, payeename)

            # title
            worksheet.write('A1', 'NON-TRADE SCHEDULE', bold)
            worksheet.write('A2', 'AS OF ' + str(dfrom) + ' to ' + str(dto), bold)

            # header
            worksheet.write('A4', 'Particulars', bold)
            worksheet.write('B4', 'Amount', bold)

            row = 5
            col = 0
            tbal = 0

            for data in result:
                if data.pcode:
                    worksheet.write(row, col, data.pname+' - '+data.pcode)
                else:
                    worksheet.write(row, col, ' NO CUSTOMER/SUPPLIER - N/A')
                worksheet.write(row, col + 1, float(format(data.balamount, '.2f')))
                row += 1
                tbal += data.balamount

            worksheet.write(row, col, 'Total')
            worksheet.write(row, col + 1, float(format(tbal, '.2f')))


        workbook.close()

        # Rewind the buffer.

        output.seek(0)

        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename=%s' % filename

        return response
    

@method_decorator(login_required, name='dispatch')
class TransExcelReport(View):
    def aging_list(self):
        return  {
            'amount_due': 0,
            'current': 0,
            'day30': 0,
            'day60': 0,
            'day90': 0,
            'day120': 0,
            'day150': 0,
            'day180': 0,
            'day210': 0,
            'over210': 0,
            'overpayment': 0,
        }
            
    def compute_grandtotal(self, datalist):
        ages = {
            'amount_due',
            'current',
            'day30',
            'day60',
            'day90',
            'day120',
            'day150',
            'day180',
            'day210',
            'over210',
            'overpayment',
        }
        grandtotal = self.aging_list()
        for cx in datalist:
            for key in ages:
                grandtotal[key] += datalist[cx]['subtotal'][key]
                
        for amount in grandtotal:
            if grandtotal[amount] < 0:
                grandtotal[amount] = '(' + "{:0,.2f}".format(abs(grandtotal[amount])) + ')'
            else:
                grandtotal[amount] = "{:0,.2f}".format(grandtotal[amount])
        # print 'grandtotal', grandtotal
        
        return grandtotal
    
    def diff_month(self, d1, d2):
        return (d1.year - d2.year) * 12 + d1.month - d2.month
    
    def get(self, request):

        dto = request.GET["dto"]
        dfrom = request.GET["dfrom"]
        transaction = request.GET["transaction"]
        chartofaccount = request.GET["chartofaccount"]
        payeecode = request.GET["payeecode"]
        payeename = request.GET["payeename"]
        report = request.GET["report"]

        output = io.BytesIO()

        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()

        # variables
        bold = workbook.add_format({'bold': 1})
        formatdate = workbook.add_format({'num_format': 'MM/DD/YYYY'})
        centertext = workbook.add_format({'bold': 1, 'align': 'center'})

        if report == '1':
            print "List of Outstanding Non-Trade Receivable"
            filename = "List of Outstanding Non-Trade Receivable.xlsx"

            result = queryLedger(dto, dfrom, transaction, chartofaccount, payeecode, payeename, isNT=True, isClosedOnly=False)

            tdebit = 0
            tcredit = 0

            # title
            worksheet.write('A1', 'LIST OF OUTSTANDING NON-TRADE RECEIVABLE', bold)
            worksheet.write('A2', 'AS OF ' + str(dfrom) + ' to ' + str(dto), bold)

            # header
            worksheet.write('A4', 'Date', bold)
            worksheet.write('B4', 'Type', bold)
            worksheet.write('C4', 'Number', bold)
            worksheet.write('D4', 'Customer', bold)
            worksheet.write('E4', 'Particulars', bold)
            worksheet.write('F4', 'Debit Amount', bold)
            worksheet.write('G4', 'Credit Amount', bold)
            worksheet.write('H4', 'Remaining Balance', bold)
            worksheet.write('I4', 'Ref Type', bold)
            worksheet.write('J4', 'Ref No', bold)
            worksheet.write('K4', 'Ref Date', bold)

            row = 5
            col = 0

            # print result

            for data in result:
                worksheet.write(row, col, data.document_date, formatdate)
                worksheet.write(row, col + 1, data.document_type)
                
                if data.orsource == 'A':
                    worksheet.write(row, col + 2, str('OR') + '' + data.document_num)
                elif data.orsource == 'C':
                    worksheet.write(row, col + 2, str('CR') + '' + data.document_num)
                else:
                    worksheet.write(row, col + 2, data.document_num)
                    
                if data.pcode:
                    worksheet.write(row, col + 3, data.pcode+'-'+data.pname)
                else:
                    worksheet.write(row, col + 3, 'N/A - NO CUSTOMER/SUPPLIER')
                    
                worksheet.write(row, col + 4, data.particulars)
                worksheet.write(row, col + 5, float(format(data.debitamount, '.2f')))
                worksheet.write(row, col + 6, float(format(data.creditamount, '.2f')))
                if data.document_refamount == 0:
                    worksheet.write(row, col + 7, '')
                elif data.document_refamount < 0:
                    worksheet.write(row, col + 7, '(' + str(data.document_refamount) + ')')
                else:
                    worksheet.write(row, col + 7, float(format(data.document_refamount, '.2f')))
                worksheet.write(row, col + 8, data.document_reftype)
                worksheet.write(row, col + 9, data.document_refnum)
                worksheet.write(row, col + 10, data.document_refdate, formatdate)

                if data.balancecode == 'D':
                    tdebit += data.debitamount
                else:
                    if data.document_refamount == 0:
                        tcredit += data.creditamount
                    else:
                        tcredit += abs(data.document_refamount)

                row += 1

            worksheet.write(row, col + 4, 'Total')
            worksheet.write(row, col + 5, float(format(tdebit, '.2f')))
            worksheet.write(row, col + 6, float(format(tcredit, '.2f')))

        elif report == '2':
            print "Aging of Non-Trade Receivable"
            filename = "Aging of Non-Trade Receivable.xlsx"

            result = queryLedgerAging(dto, dfrom, transaction, chartofaccount, payeecode, payeename)

            # title
            worksheet.write('A1', 'AGING NON-TRADE RECEIVABLE', bold)
            worksheet.write('A2', 'AS OF ' + str(dfrom) + ' to ' + str(dto), bold)

            # header
            worksheet.write('A4', 'Customer', bold)
            worksheet.write('B4', 'Number', bold)
            worksheet.write('C4', 'Date', bold)
            worksheet.write('D4', 'Amount Due', bold)
            worksheet.write('E4', 'Current', bold)
            worksheet.write('F4', '30 Days', bold)
            worksheet.write('G4', '60 Days', bold)
            worksheet.write('H4', '90 Days', bold)
            worksheet.write('I4', '120 Days', bold)
            worksheet.write('J4', '150 Days', bold)
            worksheet.write('K4', '180 Days', bold)
            worksheet.write('L4', '210 Days', bold)
            worksheet.write('M4', 'Over 210 Days', bold)
            worksheet.write('N4', 'Over-Payment', bold)

            row = 5
            col = 0

            date_to_age = dt.strptime(dto, '%Y-%m-%d').date()
            aging_brackets = [
                ("current", lambda month_past_due: month_past_due == 0),
                ("day30", lambda month_past_due: month_past_due == -1),
                ("day60", lambda month_past_due: month_past_due == -2),
                ("day90", lambda month_past_due: month_past_due == -3),
                ("day120", lambda month_past_due: month_past_due == -4),
                ("day150", lambda month_past_due: month_past_due == -5),
                ("day180", lambda month_past_due: month_past_due == -6),
                ("day210", lambda month_past_due: month_past_due == -7),
                ("over210", lambda month_past_due: month_past_due <= -8),
                ("overpayment", lambda balancecode : balancecode == 'C'),
            ]
            
            datalist = {}
            for receivable in result:
                amount = 0
                balancecode = receivable.balancecode
                aging_parameter = self.aging_list()
                
                customer_code = str(receivable.pcode) +" - "+ str(receivable.pname)
                if customer_code not in datalist:
                    amount_list = self.aging_list()
                    datalist[customer_code] = {
                        'transactions': [],
                        'subtotal': amount_list,
                    }
                    
                if balancecode == 'D':
                    amount = receivable.debitamount
                    months = self.diff_month(receivable.document_date, date_to_age)
                    
                    for bracket_name, condition in aging_brackets[:-1]:
                        if condition(months):
                            aging_parameter[bracket_name] = amount
                            datalist[customer_code]['subtotal']['amount_due'] += amount
                            datalist[customer_code]['subtotal'][bracket_name] += amount
                            break
                else:
                    if receivable.document_refamount == 0:
                        amount = receivable.creditamount
                    else:
                        amount = receivable.document_refamount
                        
                    aging_parameter['overpayment'] = amount
                    datalist[customer_code]['subtotal']['amount_due'] += -abs(amount)
                    datalist[customer_code]['subtotal']['overpayment'] += -abs(amount)
                    
                aging_parameter['amount_due'] = amount
                
                details = {
                    'main_detail': receivable,
                    'aging_detail': aging_parameter
                }
                datalist[customer_code]['transactions'].append(details)
                
            grandtotal = self.compute_grandtotal(datalist)
                
            for customer, customer_data in datalist.items():
                print 'customer', customer, customer_data['transactions']
                worksheet.write(row, col, customer)
                for item in customer_data['transactions']:
                    row += 1
                    worksheet.write(row, col + 1, str(item['main_detail'].document_type) + str(item['main_detail'].document_num))
                    worksheet.write(row, col + 2, item['main_detail'].document_date, formatdate)
                    if item['main_detail'].balancecode == 'D':
                        worksheet.write(row, col + 3, float(format(item['aging_detail']['amount_due'], '.2f')))
                    else:
                        worksheet.write(row, col + 3, '(' + str(float(format(item['aging_detail']['amount_due'], '.2f'))) + ')')
                    
                    worksheet.write(row, col + 4, item['aging_detail']['current'])
                    worksheet.write(row, col + 5, item['aging_detail']['day30'])
                    worksheet.write(row, col + 6, item['aging_detail']['day60'])
                    worksheet.write(row, col + 7, item['aging_detail']['day90'])
                    worksheet.write(row, col + 8, item['aging_detail']['day120'])
                    worksheet.write(row, col + 9, item['aging_detail']['day150'])
                    worksheet.write(row, col + 10, item['aging_detail']['day180'])
                    worksheet.write(row, col + 11, item['aging_detail']['day210'])
                    worksheet.write(row, col + 12, item['aging_detail']['over210'])
                    worksheet.write(row, col + 13, item['aging_detail']['overpayment'])
                row += 1
                worksheet.write(row, col + 1, 'Subtotal')
                if customer_data['subtotal']['amount_due'] < 0:
                    worksheet.write(row, col + 3, '('+ str(customer_data['subtotal']['amount_due']) +')')
                else:
                    worksheet.write(row, col + 3, customer_data['subtotal']['amount_due'])
                worksheet.write(row, col + 4, customer_data['subtotal']['current'])
                worksheet.write(row, col + 5, customer_data['subtotal']['day30'])
                worksheet.write(row, col + 6, customer_data['subtotal']['day60'])
                worksheet.write(row, col + 7, customer_data['subtotal']['day90'])
                worksheet.write(row, col + 8, customer_data['subtotal']['day120'])
                worksheet.write(row, col + 9, customer_data['subtotal']['day150'])
                worksheet.write(row, col + 10, customer_data['subtotal']['day180'])
                worksheet.write(row, col + 11, customer_data['subtotal']['day210'])
                worksheet.write(row, col + 12, customer_data['subtotal']['over210'])
                worksheet.write(row, col + 13, '('+ str(abs(customer_data['subtotal']['overpayment'])) +')')
            row += 1
            worksheet.write(row, col + 1, 'Grand Total')
            worksheet.write(row, col + 3, grandtotal['amount_due'])
            worksheet.write(row, col + 4, grandtotal['current'])
            worksheet.write(row, col + 5, grandtotal['day30'])
            worksheet.write(row, col + 6, grandtotal['day60'])
            worksheet.write(row, col + 7, grandtotal['day90'])
            worksheet.write(row, col + 8, grandtotal['day120'])
            worksheet.write(row, col + 9, grandtotal['day150'])
            worksheet.write(row, col + 10, grandtotal['day180'])
            worksheet.write(row, col + 11, grandtotal['day210'])
            worksheet.write(row, col + 12, grandtotal['over210'])
            worksheet.write(row, col + 13, grandtotal['overpayment'])
            
        elif report == '3':
            print 'Sales Book for Outstanding Non-Trade Receivable'
            filename = "Sales Book for Outstanding Non-Trade Receivable.xlsx"
            
            result = querySalesBookAR(dto, dfrom, transaction, chartofaccount, payeecode, payeename)

            # title
            worksheet.write('A1', 'LIST OF OUTSTANDING NON-TRADE RECEIVABLE', bold)
            worksheet.write('A2', 'AS OF ' + str(dfrom) + ' to ' + str(dto), bold)

            # header
            worksheet.write('A4', 'Date', bold)
            worksheet.write('B4', 'Type', bold)
            worksheet.write('C4', 'Number', bold)
            worksheet.write('D4', 'TIN', bold)
            worksheet.write('E4', 'Customer', bold)
            worksheet.write('F4', 'Address', bold)
            worksheet.write('G4', 'Particulars', bold)
            worksheet.write('H4', 'Amount', bold)
            worksheet.write('I4', 'Discount Amount', bold)
            worksheet.write('J4', 'VAT Amount', bold)
            worksheet.write('K4', 'NET Sales', bold)

            row = 5
            col = 0
            totalnetsales = 0
            # print result
            
            for data in result:
                worksheet.write(row, col, data.document_date, formatdate)
                worksheet.write(row, col + 1, data.document_type)
                worksheet.write(row, col + 2, data.document_num)
                worksheet.write(row, col + 3, data.ptin)
                    
                if data.pcode:
                    worksheet.write(row, col + 4, data.pcode+'-'+data.pname)
                else:
                    worksheet.write(row, col + 4, 'N/A - NO CUSTOMER/SUPPLIER')
                    
                worksheet.write(row, col + 5, str(data.paddress1) +' '+ str(data.paddress2) +' '+ str(data.paddress3))
                worksheet.write(row, col + 6, data.particulars)
                
                if data.jv_amount:
                    worksheet.write(row, col + 7, float(format(data.jv_amount, '.2f')))
                else: 
                    worksheet.write(row, col + 7, float(format(data.jv_amount, '.2f')))
                    
                worksheet.write(row, col + 8, '')
                
                if data.jv_vat:
                    worksheet.write(row, col + 9, float(format(data.jv_vat, '.2f')))
                else:
                    worksheet.write(row, col + 9, '')
                
                if data.balancecode == 'D':
                    worksheet.write(row, col + 10, float(format(data.debitamount, '.2f')))
                    totalnetsales += data.debitamount
                else:
                    if data.document_refamount == 0:
                        worksheet.write(row, col + 10, '('+ str(data.creditamount) +')')
                        totalnetsales += -abs(data.creditamount)
                    else:
                        worksheet.write(row, col + 10, '('+ str(data.document_refamount) +')')
                        totalnetsales += -abs(data.document_refamount)
                
                row += 1

            worksheet.write(row, col + 9, 'Total')
            worksheet.write(row, col + 10, float(format(totalnetsales, '.2f')))
        workbook.close()

        # Rewind the buffer.

        output.seek(0)

        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename=%s' % filename

        return response


def queryLedger(dto, dfrom, transaction, chartofaccount, payeecode, payeename, isNT, isClosedOnly):

    conchart = ""
    orderby = "ORDER BY document_date ASC, FIELD(a.document_type, 'AP','CV','JV','OR')"
    conpayeecode = ""
    conpayeename = ""
    conunapplied = ""
    conclosedonly = ""

    if chartofaccount:
        conchart = "SELECT id FROM chartofaccount WHERE id = '"+str(chartofaccount)+"' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode"
    else:
        conchart  = "SELECT id FROM chartofaccount WHERE main = '"+str(transaction)+"' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode"

    if payeecode:
        conpayeecode = "AND IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code, IF (b.setup_customer != '', scust.code, ssup.code))) = '"+str(payeecode)+"'"

    if payeename:
        conpayeename = "AND IF (b.customer_enable = 'Y', dcust.name, IF (b.supplier_enable = 'Y', dsup.name, IF (b.setup_customer != '', scust.name, ssup.name))) LIKE '%"+str(payeename)+"%'"
        
    if isNT:
        conunapplied = "AND (is_closed = 0 OR is_closed IS NULL)"
        
    if isClosedOnly:
        conclosedonly = "AND is_closed = 1"
        orderby = "ORDER BY tag_id, document_date ASC"

    print conchart
    ''' Create query '''
    cursor = connection.cursor()

    query = "SELECT a.id, a.document_type, a.document_id, a.document_num, a.document_date, a.subtype, a.particulars, a.amount, a.chartofaccount_id, a.document_reftype, a.document_refnum, a.document_refdate, a.document_refamount, a.document_amount, a.is_closed, a.tag_id, " \
            "a.balancecode, IF (a.balancecode = 'C', a.amount, 0) AS creditamount, IF (a.balancecode = 'D', a.amount, 0) AS debitamount, " \
            "a.document_customer_id, a.document_supplier_id,  " \
            "b.accountcode, b.description, b.customer_enable, b.supplier_enable, b.nontrade, b.setup_customer, b.setup_supplier, " \
            "IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code, IF (b.setup_customer != '', scust.code, ssup.code))) AS pcode,  " \
            "IF (b.customer_enable = 'Y', dcust.name, IF (b.supplier_enable = 'Y', dsup.name, IF (b.setup_customer != '', scust.name, ssup.name))) AS pname, " \
            "IF (b.customer_enable = 'Y', dcust.tin, IF (b.supplier_enable = 'Y', dsup.tin, IF (b.setup_customer != '', scust.tin, ssup.tin))) AS ptin, om.orsource " \
            "FROM subledger AS a " \
            "LEFT OUTER JOIN chartofaccount AS b ON b.id = a.chartofaccount_id " \
            "LEFT OUTER JOIN customer AS dcust ON dcust.id = a.customer_id " \
            "LEFT OUTER JOIN customer AS scust ON scust.id =  b.setup_customer " \
            "LEFT OUTER JOIN supplier AS dsup ON dsup.id = a.document_supplier_id " \
            "LEFT OUTER JOIN supplier AS ssup ON ssup.id = b.setup_supplier " \
            "left outer join ordetail as od on (od.id = a.document_id and a.document_type = 'OR') " \
            "left outer join ormain as om on om.id = od.ormain_id " \
            "WHERE a.chartofaccount_id IN ("+str(conchart)+") " \
            ""+str(conpayeecode)+" "+str(conpayeename)+" "+str(conunapplied)+" "+str(conclosedonly)+" AND DATE(document_date) >= '"+str(dfrom)+"' AND DATE(document_date) <= '"+str(dto)+"' "+str(orderby)

    ##"LEFT OUTER JOIN customer AS dcust ON dcust.id = a.document_customer_id "
    # print query
    # print '****'

    cursor.execute(query)
    result = namedtuplefetchall(cursor)
    
    cursor.close()
    
    return result


def queryScheduled(dto, dfrom, transaction, chartofaccount, payeecode, payeename):
    chart = Chartofaccount.objects.filter(isdeleted=0, accounttype='P', nontrade='Y', id=chartofaccount).first()

    conchart = ""
    orderby = "ORDER BY document_date ASC, FIELD(a.document_type, 'AP','CV','JV','OR')"
    #orderby = "ORDER BY IFNULL(z.pcode, ' N/A') ASC, ' NO CUSTOMER/SUPPLIER') ASC"
    conpayeecode = ""
    conpayeename = ""
    conbeg = ""
    contran = ""

    if transaction == '1':
        #conbeg = "LEFT OUTER JOIN beginningbalance AS beg ON (beg.code_id = z.pid AND beg.trans_code = 'NT' AND trans_type = 'AR' AND beg.accountcode = '"+str(chart.accountcode)+"' )"
        print 'hy'
        conbeg = "WHERE b.accountcode = '" + str(chart.accountcode) + "' AND b.trans_code = 'NT' AND b.trans_type = 'AR'"

    else:
        #conbeg = "LEFT OUTER JOIN beginningbalance AS beg ON (beg.code_id = z.pid AND beg.trans_code = 'NT' AND trans_type = 'AP' AND beg.accountcode = '"+str(chart.accountcode)+"' )"
        conbeg = "WHERE b.accountcode = '"+str(chart.accountcode)+"' AND b.trans_code = 'NT' AND b.trans_type = 'AP'"
        print 'hy'

    contran = "LEFT OUTER JOIN ( " \
                " SELECT a.id, a.document_type, a.document_id, a.document_num, a.document_date, a.subtype, a.particulars, a.chartofaccount_id, a.document_reftype, " \
                " a.document_refnum, a.document_refdate, a.balancecode, " \
                " SUM(IF (a.balancecode = 'C', a.amount, 0)) AS creditamount, " \
                " SUM(IF (a.balancecode = 'D', a.amount, 0)) AS debitamount, " \
                " (SUM(IF (a.balancecode = 'D', a.amount, 0)) - SUM(IF (a.balancecode = 'C', a.amount, 0))) AS bal,  " \
                " a.document_customer_id, a.document_supplier_id,  b.accountcode, b.description, b.customer_enable, b.supplier_enable, b.nontrade, b.setup_customer, b.setup_supplier, " \
                " IF (b.customer_enable = 'Y', dcust.id, IF (b.supplier_enable = 'Y', dsup.id, IF (b.setup_customer != '', scust.id, ssup.id))) AS pid, " \
                " IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code, " \
                " IF (b.setup_customer != '', scust.code, ssup.code))) AS pcode,  " \
                " IF (b.customer_enable = 'Y', dcust.name, IF (b.supplier_enable = 'Y', dsup.name, IF (b.setup_customer != '', scust.name, ssup.name))) AS pname, " \
                " IF (b.customer_enable = 'Y', dcust.tin, IF (b.supplier_enable = 'Y', dsup.tin, IF (b.setup_customer != '', scust.tin, ssup.tin))) AS ptin " \
                " FROM subledger AS a " \
                " LEFT OUTER JOIN chartofaccount AS b ON b.id = a.chartofaccount_id " \
                " LEFT OUTER JOIN customer AS dcust ON dcust.id = a.document_customer_id " \
                " LEFT OUTER JOIN customer AS scust ON scust.id = b.setup_customer " \
                " LEFT OUTER JOIN supplier AS dsup ON dsup.id = a.document_supplier_id " \
                " LEFT OUTER JOIN supplier AS ssup ON ssup.id = b.setup_supplier " \
                " WHERE a.chartofaccount_id IN ( " \
                "     SELECT id FROM chartofaccount WHERE id = '" + str(chartofaccount) + "' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode) " \
                            "     AND DATE(a.document_date) > '2019-01-01' AND DATE(a.document_date) < '" + str(dfrom) + "' " \
                    "     GROUP BY IF (b.customer_enable = 'Y', dcust.id, IF (b.supplier_enable = 'Y', dsup.id, IF (b.setup_customer != '', scust.id, ssup.id))) " \
                    ") AS extrans ON extrans.pid = z.pid"

    if chartofaccount:
        conchart = "SELECT id FROM chartofaccount WHERE id = '"+str(chartofaccount)+"' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode"
    else:
        conchart  = "SELECT id FROM chartofaccount WHERE main = '"+str(transaction)+"' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode"

    if payeecode:
        conpayeecode = "AND IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code, IF (b.setup_customer != '', scust.code, ssup.code))) = '"+str(payeecode)+"'"

    if payeename:
        conpayeename = "AND IF (b.customer_enable = 'Y', dcust.name, IF (b.supplier_enable = 'Y', dsup.name, IF (b.setup_customer != '', scust.name, ssup.name))) LIKE '%"+str(payeename)+"%'"

    print conchart
    ''' Create query '''
    cursor = connection.cursor()

    query = "SELECT zz.pcode, zz.pname, pid, SUM(zz.creditamount) AS creditamount, SUM(zz.debitamount) AS debitamount, IFNULL(zz.beg_code, 'D') AS beg_code, " \
            "SUM(IFNULL(zz.begamt, 0)) AS begamt, SUM(zz.balamountx) AS balamountx, SUM(zz.extrabal) AS extrabal, SUM(zz.balamount) AS balamount " \
            "FROM (" \
            "SELECT IFNULL(z.pcode, 'NA') AS pcode, IFNULL(z.pname, ' NO CUSTOMER/SUPPLIER') AS pname, z.pid, " \
            "SUM(z.creditamount) AS creditamount, SUM(z.debitamount) AS debitamount, " \
            "IFNULL(z.balancecode, 'D') AS beg_code, 0 AS begamt, " \
            "(SUM(IF(z.balancecode = 'D', z.debitamount, 0)) - SUM(IF(z.balancecode = 'C', z.creditamount, 0))) AS balamountx, 0 AS extrabal, " \
            "(SUM(IF(z.balancecode = 'D', z.debitamount, 0)) - SUM(IF(z.balancecode = 'C', z.creditamount, 0))) AS balamount " \
            "FROM ( SELECT a.id, a.document_type, a.document_id, a.document_num, a.document_date, a.subtype, a.particulars, a.chartofaccount_id, a.document_reftype, a.document_refnum, a.document_refdate, " \
            "a.balancecode, IF (a.balancecode = 'C', a.amount, 0) AS creditamount, IF (a.balancecode = 'D', a.amount, 0) AS debitamount, " \
            "a.document_customer_id, a.document_supplier_id, '"+str(chart.balancecode)+"' AS cbcode, " \
            "b.accountcode, b.description, b.customer_enable, b.supplier_enable, b.nontrade, b.setup_customer, b.setup_supplier, " \
            "IF (b.customer_enable = 'Y', dcust.id, IF (b.supplier_enable = 'Y', dsup.id, IF (b.setup_customer != '', scust.id, ssup.id))) AS pid,  " \
            "IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code, IF (b.setup_customer != '', scust.code, ssup.code))) AS pcode,  " \
            "IF (b.customer_enable = 'Y', dcust.name, IF (b.supplier_enable = 'Y', dsup.name, IF (b.setup_customer != '', scust.name, ssup.name))) AS pname, " \
            "IF (b.customer_enable = 'Y', dcust.tin, IF (b.supplier_enable = 'Y', dsup.tin, IF (b.setup_customer != '', scust.tin, ssup.tin))) AS ptin " \
            "FROM subledger AS a " \
            "LEFT OUTER JOIN chartofaccount AS b ON b.id = a.chartofaccount_id " \
            "LEFT OUTER JOIN customer AS dcust ON dcust.id = a.customer_id " \
            "LEFT OUTER JOIN customer AS scust ON scust.id =  b.setup_customer " \
            "LEFT OUTER JOIN supplier AS dsup ON dsup.id = a.document_supplier_id " \
            "LEFT OUTER JOIN supplier AS ssup ON ssup.id = b.setup_supplier " \
            "WHERE a.chartofaccount_id IN ("+str(conchart)+") " \
            ""+str(conpayeecode)+" "+str(conpayeename)+" AND DATE(document_date) >= '"+str(dfrom)+"' AND DATE(document_date) <= '"+str(dto)+"' "+str(orderby) +") AS z GROUP BY z.pcode " \
            "UNION " \
            "SELECT IFNULL(zx.pcode, 'NA') AS pcode, IFNULL(zx.pname, ' NO CUSTOMER/SUPPLIER') AS pname, zx.pid, " \
            "SUM(zx.creditamount) AS creditamount, SUM(zx.debitamount) AS debitamount, " \
            "IFNULL(zx.balancecode, 'D') AS beg_code, 0 AS begamt, " \
            "0 AS balamountx, " \
            "zx.bal  AS extrabal, " \
            "zx.bal AS balamount " \
            "FROM (  " \
            "SELECT a.id, a.document_type, a.document_id, a.document_num, a.document_date, a.subtype, a.particulars, a.chartofaccount_id, a.document_reftype,  " \
            "a.document_refnum, a.document_refdate, a.balancecode,  " \
            "SUM(IF (a.balancecode = 'C', a.amount, 0)) AS creditamount, " \
            "SUM(IF (a.balancecode = 'D', a.amount, 0)) AS debitamount,  " \
            "(SUM(IF (a.balancecode = 'D', a.amount, 0)) - SUM(IF (a.balancecode = 'C', a.amount, 0))) AS bal,  " \
            "a.document_customer_id, a.document_supplier_id,  b.accountcode, b.description, b.customer_enable, b.supplier_enable, b.nontrade, b.setup_customer, b.setup_supplier,  " \
            "IF (b.customer_enable = 'Y', dcust.id, IF (b.supplier_enable = 'Y', dsup.id, IF (b.setup_customer != '', scust.id, ssup.id))) AS pid, " \
            "IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code,  " \
            "IF (b.setup_customer != '', scust.code, ssup.code))) AS pcode, " \
            "IF (b.customer_enable = 'Y', dcust.name, IF (b.supplier_enable = 'Y', dsup.name, IF (b.setup_customer != '', scust.name, ssup.name))) AS pname, " \
            "IF (b.customer_enable = 'Y', dcust.tin, IF (b.supplier_enable = 'Y', dsup.tin, IF (b.setup_customer != '', scust.tin, ssup.tin))) AS ptin " \
            "FROM subledger AS a " \
            "LEFT OUTER JOIN chartofaccount AS b ON b.id = a.chartofaccount_id " \
            "LEFT OUTER JOIN customer AS dcust ON dcust.id = a.customer_id " \
            "LEFT OUTER JOIN customer AS scust ON scust.id = b.setup_customer " \
            "LEFT OUTER JOIN supplier AS dsup ON dsup.id = a.document_supplier_id " \
            "LEFT OUTER JOIN supplier AS ssup ON ssup.id = b.setup_supplier " \
            "WHERE a.chartofaccount_id IN ("+str(conchart)+") " \
            "AND DATE(a.document_date) > '2019-01-01' AND DATE(a.document_date) <= '"+str(dfrom)+"' " \
            "GROUP BY IF (b.customer_enable = 'Y', dcust.id, IF (b.supplier_enable = 'Y', dsup.id, IF (b.setup_customer != '', scust.id, ssup.id)))) AS zx " \
            "GROUP BY zx.pcode " \
            "UNION " \
            "SELECT IFNULL(s.code, 'N/A') AS pcode, IFNULL(s.name, ' NO CUSTOMER/SUPPLIER') AS pname, s.id AS pid, " \
            "0 AS creditamount, 0 AS debitamount, " \
            "IFNULL(b.beg_code, 'D') AS beg_code, IF (b.beg_code = 'D', IFNULL(b.beg_amt, 0), IFNULL(b.beg_amt, 0) * -1) AS begamt,  " \
            "0 AS balamountx, 0 AS extrabal, IF (b.beg_code = 'D', IFNULL(b.beg_amt, 0), IFNULL(b.beg_amt, 0) * -1) AS balamount  " \
            "FROM beginningbalance  AS b " \
            "LEFT OUTER JOIN supplier AS s ON s.id = b.code_id " \
            " "+str(conbeg)+") zz " \
            "GROUP BY zz.pid ORDER BY zz.pname"

    cursor.execute(query)
    result = namedtuplefetchall(cursor)
    
    cursor.close()
    
    return result


def querySOA(dto, dfrom, transaction, chartofaccount, payeecode, payeename):

    conchart = ""
    orderby = "ORDER BY document_date ASC, FIELD(a.document_type, 'AP','CV','JV','OR')"
    conpayeecode = ""
    conpayeename = ""

    if chartofaccount:
        conchart = "SELECT id FROM chartofaccount WHERE id = '"+str(chartofaccount)+"' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode"
    else:
        conchart  = "SELECT id FROM chartofaccount WHERE main = '"+str(transaction)+"' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode"

    if payeecode:
        conpayeecode = "AND IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code, IF (b.setup_customer != '', scust.code, ssup.code))) = '"+str(payeecode)+"'"

    if payeename:
        conpayeename = "AND IF (b.customer_enable = 'Y', dcust.name, IF (b.supplier_enable = 'Y', dsup.name, IF (b.setup_customer != '', scust.name, ssup.name))) = '"+str(payeename)+"'"

    print conchart
    ''' Create query '''
    cursor = connection.cursor()

    query = "SELECT a.id, a.document_type, a.document_id, a.document_num, a.document_date, a.subtype, a.particulars, a.chartofaccount_id, a.document_reftype, a.document_refnum, a.document_refdate, " \
            "a.balancecode, IF (a.balancecode = 'C', a.amount, 0) AS creditamount, IF (a.balancecode = 'D', a.amount, 0) AS debitamount, " \
            "a.document_customer_id, a.document_supplier_id,  " \
            "b.accountcode, b.description, b.customer_enable, b.supplier_enable, b.nontrade, b.setup_customer, b.setup_supplier, " \
            "IF (b.customer_enable = 'Y', dcust.id, IF (b.supplier_enable = 'Y', dsup.id, IF (b.setup_customer != '', scust.id, ssup.id))) AS pid,  " \
            "IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code, IF (b.setup_customer != '', scust.code, ssup.code))) AS pcode,  " \
            "IF (b.customer_enable = 'Y', dcust.name, IF (b.supplier_enable = 'Y', dsup.name, IF (b.setup_customer != '', scust.name, ssup.name))) AS pname, " \
            "IF (b.customer_enable = 'Y', dcust.tin, IF (b.supplier_enable = 'Y', dsup.tin, IF (b.setup_customer != '', scust.tin, ssup.tin))) AS ptin " \
            "FROM subledger AS a " \
            "LEFT OUTER JOIN chartofaccount AS b ON b.id = a.chartofaccount_id " \
            "LEFT OUTER JOIN customer AS dcust ON dcust.id = a.customer_id " \
            "LEFT OUTER JOIN customer AS scust ON scust.id =  b.setup_customer " \
            "LEFT OUTER JOIN supplier AS dsup ON dsup.id = a.document_supplier_id " \
            "LEFT OUTER JOIN supplier AS ssup ON ssup.id = b.setup_supplier " \
            "WHERE a.chartofaccount_id IN ("+str(conchart)+") " \
            ""+str(conpayeecode)+" "+str(conpayeename)+" AND DATE(document_date) >= '"+str(dfrom)+"' AND DATE(document_date) <= '"+str(dto)+"' "+str(orderby)

    #print query

    cursor.execute(query)
    result = namedtuplefetchall(cursor)
    
    cursor.close()

    return result


def querySOATran(dto, dfrom, transaction, chartofaccount, payeecode, payeename):

    conchart = ""
    orderby = "ORDER BY document_date ASC, FIELD(a.document_type, 'AP','CV','JV','OR')"
    conpayeecode = ""
    conpayeename = ""

    if chartofaccount:
        conchart = "SELECT id FROM chartofaccount WHERE id = '"+str(chartofaccount)+"' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode"
    else:
        conchart  = "SELECT id FROM chartofaccount WHERE main = '"+str(transaction)+"' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode"

    if payeecode:
        conpayeecode = "AND IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code, IF (b.setup_customer != '', scust.code, ssup.code))) = '"+str(payeecode)+"'"

    if payeename:
        conpayeename = "AND IF (b.customer_enable = 'Y', dcust.name, IF (b.supplier_enable = 'Y', dsup.name, IF (b.setup_customer != '', scust.name, ssup.name))) = '"+str(payeename)+"'"

    print conchart
    ''' Create query '''
    cursor = connection.cursor()

    query = "SELECT a.id, a.document_type, a.document_id, a.document_num, a.document_date, a.subtype, a.particulars, a.chartofaccount_id, a.document_reftype, a.document_refnum, a.document_refdate, " \
            "a.balancecode, IF (a.balancecode = 'C', a.amount, 0) AS creditamount, IF (a.balancecode = 'D', a.amount, 0) AS debitamount, " \
            "a.document_customer_id, a.document_supplier_id,  " \
            "b.accountcode, b.description, b.customer_enable, b.supplier_enable, b.nontrade, b.setup_customer, b.setup_supplier, " \
            "IF (b.customer_enable = 'Y', dcust.id, IF (b.supplier_enable = 'Y', dsup.id, IF (b.setup_customer != '', scust.id, ssup.id))) AS pid,  " \
            "IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code, IF (b.setup_customer != '', scust.code, ssup.code))) AS pcode,  " \
            "IF (b.customer_enable = 'Y', dcust.name, IF (b.supplier_enable = 'Y', dsup.name, IF (b.setup_customer != '', scust.name, ssup.name))) AS pname, " \
            "IF (b.customer_enable = 'Y', dcust.tin, IF (b.supplier_enable = 'Y', dsup.tin, IF (b.setup_customer != '', scust.tin, ssup.tin))) AS ptin " \
            "FROM subledger AS a " \
            "LEFT OUTER JOIN chartofaccount AS b ON b.id = a.chartofaccount_id " \
            "LEFT OUTER JOIN customer AS dcust ON dcust.id = a.customer_id " \
            "LEFT OUTER JOIN customer AS scust ON scust.id =  b.setup_customer " \
            "LEFT OUTER JOIN supplier AS dsup ON dsup.id = a.document_supplier_id " \
            "LEFT OUTER JOIN supplier AS ssup ON ssup.id = b.setup_supplier " \
            "WHERE a.chartofaccount_id IN ("+str(conchart)+") " \
            ""+str(conpayeecode)+" "+str(conpayeename)+" AND DATE(document_date) > '"+str(dfrom)+"' AND DATE(document_date) < '"+str(dto)+"' "+str(orderby)

    print 'querySOATran'
    # print query

    cursor.execute(query)
    result = namedtuplefetchall(cursor)
    
    cursor.close()
    
    return result


def begBalance(id):

    ''' Create query '''
    cursor = connection.cursor()

    query = "SELECT * FROM beginningbalance WHERE trans_code = 'NT' AND code_id = '"+str(id)+"'"

    # print query

    cursor.execute(query)
    result = namedtuplefetchall(cursor)
    
    cursor.close()
    
    return result


def queryTaggedARNonTrade(dto, dfrom):
    ''' Create query '''
    cursor = connection.cursor()

    query = "SELECT a.id, a.document_type, a.document_id, a.document_num, a.document_date, a.subtype, a.particulars, a.chartofaccount_id, a.document_reftype, a.document_refnum, a.document_refdate, a.document_refamount, a.amount, "\
            "a.balancecode, a.document_customer_id, a.document_supplier_id, a.status, a.tag_id, sup.code, sup.name, " \
            "IF (a.balancecode = 'C', a.amount, 0) AS creditamount, " \
            "IF (a.balancecode = 'D', a.amount, 0) AS debitamount " \
            "FROM subledger AS a " \
            "LEFT OUTER JOIN supplier AS sup ON a.document_supplier_id = sup.id " \
            "WHERE a.chartofaccount_id IN (SELECT id FROM chartofaccount WHERE main = '2' AND isdeleted=0 AND accounttype='P' AND nontrade = 'Y' ORDER BY accountcode) " \
            "AND a.document_refnum IS NOT NULL "\
            "AND a.status <> 'C' " \
            "AND DATE(document_date) >= '"+str(dfrom)+"' " \
            "AND DATE(document_date) <= '"+str(dto)+"' " \
            "ORDER BY sup.name ASC, a.tag_id, document_date ASC, a.amount DESC, FIELD(a.document_type, 'AP','CV','JV','OR')"
    
    cursor.execute(query)
    result = namedtuplefetchall(cursor)
    
    cursor.close()
    
    return result


def queryLedgerAging(dto, dfrom, transaction, chartofaccount, payeecode, payeename):
    # time benchmarking
    # start_time = dt.now()
    conchart = ""
    orderby = "ORDER BY document_date ASC, FIELD(a.document_type, 'AP','CV','JV','OR')"
    conpayeecode = ""
    conpayeename = ""

    if chartofaccount:
        conchart = "SELECT id FROM chartofaccount WHERE id = '"+str(chartofaccount)+"' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode"
    else:
        conchart  = "SELECT id FROM chartofaccount WHERE main = '"+str(transaction)+"' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode"

    if payeecode:
        conpayeecode = "AND IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code, IF (b.setup_customer != '', scust.code, ssup.code))) = '"+str(payeecode)+"'"

    if payeename:
        conpayeename = "AND IF (b.customer_enable = 'Y', dcust.name, IF (b.supplier_enable = 'Y', dsup.name, IF (b.setup_customer != '', scust.name, ssup.name))) LIKE '%"+str(payeename)+"%'"

    print conchart
    ''' Create query '''
    cursor = connection.cursor()

    query = "SELECT a.id, a.document_type, a.document_id, a.document_num, a.document_date, a.subtype, a.amount, a.chartofaccount_id, a.document_refamount, a.is_closed, " \
            "a.balancecode, IF (a.balancecode = 'C', a.amount, 0) AS creditamount, IF (a.balancecode = 'D', a.amount, 0) AS debitamount, " \
            "a.document_customer_id, a.document_supplier_id,  " \
            "b.accountcode, b.description, b.customer_enable, b.supplier_enable, b.nontrade, b.setup_customer, b.setup_supplier, " \
            "IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code, IF (b.setup_customer != '', scust.code, ssup.code))) AS pcode,  " \
            "IF (b.customer_enable = 'Y', dcust.name, IF (b.supplier_enable = 'Y', dsup.name, IF (b.setup_customer != '', scust.name, ssup.name))) AS pname, " \
            "IF (b.customer_enable = 'Y', dcust.tin, IF (b.supplier_enable = 'Y', dsup.tin, IF (b.setup_customer != '', scust.tin, ssup.tin))) AS ptin, om.orsource " \
            "FROM subledger AS a " \
            "LEFT OUTER JOIN chartofaccount AS b ON b.id = a.chartofaccount_id " \
            "LEFT OUTER JOIN customer AS dcust ON dcust.id = a.customer_id " \
            "LEFT OUTER JOIN customer AS scust ON scust.id = b.setup_customer " \
            "LEFT OUTER JOIN supplier AS dsup ON dsup.id = a.document_supplier_id " \
            "LEFT OUTER JOIN supplier AS ssup ON ssup.id = b.setup_supplier " \
            "left outer join ordetail as od on (od.id = a.document_id and a.document_type = 'OR') " \
            "left outer join ormain as om on om.id = od.ormain_id " \
            "WHERE a.chartofaccount_id IN ("+str(conchart)+") " \
            ""+str(conpayeecode)+" "+str(conpayeename)+" AND (is_closed = 0 OR is_closed IS NULL) AND DATE(document_date) >= '"+str(dfrom)+"' AND DATE(document_date) <= '"+str(dto)+"' "+str(orderby)

    # print query
    # print '****'

    cursor.execute(query)
    result = namedtuplefetchall(cursor)
    
    cursor.close()
    # end_time = dt.now()
    # time_difference = end_time - start_time
    # print 'time', time_difference.total_seconds()
    return result
    
    
def querySalesBookAR(dto, dfrom, transaction, chartofaccount, payeecode, payeename):
    # time benchmarking
    # start_time = dt.now()
    conchart = ""
    orderby = "ORDER BY document_date ASC, FIELD(a.document_type, 'AP','CV','JV','OR')"
    conpayeecode = ""
    conpayeename = ""

    if chartofaccount:
        conchart = "SELECT id FROM chartofaccount WHERE id = '"+str(chartofaccount)+"' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode"
    else:
        conchart  = "SELECT id FROM chartofaccount WHERE main = '"+str(transaction)+"' AND isdeleted=0 AND accounttype='P' AND nontrade='Y' ORDER BY accountcode"

    if payeecode:
        conpayeecode = "AND IF (b.customer_enable = 'Y', dcust.code, IF (b.supplier_enable = 'Y', dsup.code, IF (b.setup_customer != '', scust.code, ssup.code))) = '"+str(payeecode)+"'"

    if payeename:
        conpayeename = "AND IF (b.customer_enable = 'Y', dcust.name, IF (b.setup_customer != '', scust.name, ssup.name)) LIKE '%"+str(payeename)+"%'"

    print conchart
    ''' Create query '''
    cursor = connection.cursor()

    query = "SELECT a.id, a.document_type, a.document_id, a.document_num, a.document_date, a.subtype, a.amount, a.chartofaccount_id, a.document_refamount, a.particulars, a.is_closed, " \
            "a.balancecode, IF (a.balancecode = 'C', a.amount, 0) AS creditamount, IF (a.balancecode = 'D', a.amount, 0) AS debitamount, " \
            "a.document_customer_id, a.document_supplier_id,  " \
            "b.accountcode, b.description, b.customer_enable, b.supplier_enable, b.nontrade, b.setup_customer, b.setup_supplier, " \
            "IF (b.customer_enable = 'Y', dcust.code, IF (b.setup_customer != '', scust.code, ssup.code)) AS pcode,  " \
            "IF (b.customer_enable = 'Y', dcust.name, IF (b.setup_customer != '', scust.name, ssup.name)) AS pname, " \
            "IF (b.customer_enable = 'Y', dcust.address1, IF (b.setup_customer != '', scust.address1, ssup.address1)) AS paddress1, " \
            "IF (b.customer_enable = 'Y', dcust.address2, IF (b.setup_customer != '', scust.address2, ssup.address2)) AS paddress2, " \
            "IF (b.customer_enable = 'Y', dcust.address3, IF (b.setup_customer != '', scust.address3, ssup.address3)) AS paddress3, " \
            "IF (b.customer_enable = 'Y', dcust.tin, IF (b.setup_customer != '', scust.tin, ssup.tin)) AS ptin, om.orsource, " \
            "IF (a.document_type = 'JV', IF ((SELECT COUNT(*) FROM jvdetail AS jd WHERE jd.jvmain_id = jvm.id HAVING (COUNT(*)=3)), (SELECT creditamount FROM jvdetail AS jd WHERE jd.jvmain_id = jvm.id AND jd.chartofaccount_id = 320), 0), 0) AS jv_vat, " \
            "IF ((SELECT jv_vat) > 0, a.amount - (SELECT jv_vat), 0) AS jv_amount " \
            "FROM subledger AS a " \
            "LEFT OUTER JOIN chartofaccount AS b ON b.id = a.chartofaccount_id " \
            "LEFT OUTER JOIN customer AS dcust ON dcust.id = a.customer_id " \
            "LEFT OUTER JOIN customer AS scust ON scust.id = b.setup_customer " \
            "LEFT OUTER JOIN supplier AS dsup ON dsup.id = a.document_supplier_id " \
            "LEFT OUTER JOIN supplier AS ssup ON ssup.id = b.setup_supplier " \
            "LEFT OUTER JOIN ordetail AS od on (od.id = a.document_id and a.document_type = 'OR') " \
            "LEFT OUTER JOIN ormain AS om on om.id = od.ormain_id " \
            "LEFT OUTER JOIN jvdetail AS jvd on (jvd.id = a.document_id AND a.document_type = 'JV') " \
            "LEFT OUTER JOIN jvmain AS jvm on jvm.id = jvd.jvmain_id " \
            "WHERE a.chartofaccount_id IN ("+str(conchart)+") " \
            ""+str(conpayeecode)+" "+str(conpayeename)+" AND (is_closed = 0 OR is_closed IS NULL) AND DATE(document_date) >= '"+str(dfrom)+"' AND DATE(document_date) <= '"+str(dto)+"' "+str(orderby)

    # print 'query', query
    # print '****'

    cursor.execute(query)
    result = namedtuplefetchall(cursor)
    
    cursor.close()
    # end_time = dt.now()
    # time_difference = end_time - start_time
    # print 'time', time_difference.total_seconds()
    # print 'result', result
    return result


def namedtuplefetchall(cursor):
    "Return all rows from a cursor as a namedtuple"
    desc = cursor.description
    nt_result = namedtuple('Result', [col[0] for col in desc])
    return [nt_result(*row) for row in cursor.fetchall()]


@csrf_exempt
def datafix(request):
    print 'AP datafix'

    chart = Chartofaccount.objects.filter(isdeleted=0).filter(customer_enable='Y').values_list('id', flat=True)
    print chart
    detail = Apdetail.objects.filter(isdeleted=0).filter(chartofaccount_id__in=chart)

    #print detail.count()
    counter = 1
    for d in detail:

        s = Subledger.objects.filter(document_type='AP').filter(document_id=d.id).first()

        if s:
            if s.document_customer_id != d.customer_id:
                print str(d.id)+'|'+str(d.customer_id)+'|'+str(d.ap_num)+'|'+str(s.document_customer_id)+'|'+str(s.document_num)
                print counter
                counter += 1

    print counter

    return 'hey'


def calculate_aging(document_date, date_to_age):
    days_past_due = (date_to_age - document_date).months
    # print 'days_past_due', days_past_due
    return days_past_due


@method_decorator(login_required, name='dispatch')
class ReportView(ListView):
    model = Subledger
    template_name = 'nontrade/report/index.html'
    
    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('nontrade.view_nontrade'):
            raise Http404
        return super(ListView, self).dispatch(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        context = super(ListView, self).get_context_data(**kwargs)
        context['chart'] = Chartofaccount.objects.filter(isdeleted=0, accounttype='P', nontrade='Y').order_by('accountcode')
        context['customer'] = Customer.objects.filter(isdeleted=0).order_by('code')

        return context
    
    
@method_decorator(login_required, name='dispatch')
class GenerateReportPDF(View):
    
    def aging_list(self):
        return  {
            'amount_due': 0,
            'current': 0,
            'day30': 0,
            'day60': 0,
            'day90': 0,
            'day120': 0,
            'day150': 0,
            'day180': 0,
            'day210': 0,
            'over210': 0,
            'overpayment': 0,
        }
        
    def compute_grandtotal(self, datalist):
        ages = {
            'amount_due',
            'current',
            'day30',
            'day60',
            'day90',
            'day120',
            'day150',
            'day180',
            'day210',
            'over210',
            'overpayment',
        }
        grandtotal = self.aging_list()
        for cx in datalist:
            for key in ages:
                grandtotal[key] += datalist[cx]['subtotal'][key]
                
        for amount in grandtotal:
            if grandtotal[amount] < 0:
                grandtotal[amount] = '(' + "{:0,.2f}".format(abs(grandtotal[amount])) + ')'
            else:
                grandtotal[amount] = "{:0,.2f}".format(grandtotal[amount])
        # print 'grandtotal', grandtotal
        
        return grandtotal
    
    def diff_month(self, d1, d2):
        return (d1.year - d2.year) * 12 + d1.month - d2.month
        
    def get(self, request):
        company = Companyparameter.objects.all().first()
        context = []
        data = []
        report = request.GET['report']
        dfrom = request.GET['from']
        dto = request.GET['to']
        chartofaccount = request.GET['chartofaccount']
        # classification = request.GET['classification']
        transaction = request.GET['transaction']
        payeecode = request.GET['customercode']
        payeename = request.GET['customername']
        title = "Non-Trade Report List"
        
        tdebit = 0
        tcredit = 0
        totalnetsales = 0
        grandtotal = {}
        
        if report == '1':
            title = "List of Outstanding Non-Trade Receivable"
            data = queryLedger(dto, dfrom, transaction, chartofaccount, payeecode, payeename, isNT=True, isClosedOnly=False)
            
        elif report == '2':
            title = 'Aging of Non-Trade Receivable'
            receivables = queryLedgerAging(dto, dfrom, transaction, chartofaccount, payeecode, payeename)
            
        elif report == '3':
            title = "Sales Book for Outstanding Non-Trade Receivable"
            data = querySalesBookAR(dto, dfrom, transaction, chartofaccount, payeecode, payeename)
        
        if report == '2':
            # today = dt.now().date()
            date_to_age = dt.strptime(dto, '%Y-%m-%d').date()
            
            # aging_brackets = [
            #     ("current", lambda days_past_due: days_past_due < 0),
            #     ("day30", lambda days_past_due: 0 <= days_past_due <= 30),
            #     ("day60", lambda days_past_due: 31 <= days_past_due <= 60),
            #     ("day90", lambda days_past_due: 61 <= days_past_due <= 90),
            #     ("day120", lambda days_past_due: 91 <= days_past_due <= 120),
            #     ("day150", lambda days_past_due: 121 <= days_past_due <= 150),
            #     ("day180", lambda days_past_due: 151 <= days_past_due <= 180),
            #     ("day210", lambda days_past_due: 181 <= days_past_due <= 210),
            #     ("over210", lambda days_past_due: days_past_due > 210),
            #     ("overpayment", lambda balancecode : balancecode == 'C'),
            # ]
            
            aging_brackets = [
                ("current", lambda month_past_due: month_past_due == 0),
                ("day30", lambda month_past_due: month_past_due == -1),
                ("day60", lambda month_past_due: month_past_due == -2),
                ("day90", lambda month_past_due: month_past_due == -3),
                ("day120", lambda month_past_due: month_past_due == -4),
                ("day150", lambda month_past_due: month_past_due == -5),
                ("day180", lambda month_past_due: month_past_due == -6),
                ("day210", lambda month_past_due: month_past_due == -7),
                ("over210", lambda month_past_due: month_past_due <= -8),
                ("overpayment", lambda balancecode : balancecode == 'C'),
            ]
            
            datalist = {}
            for receivable in receivables:
                
                amount = 0
                balancecode = receivable.balancecode
                aging_parameter = self.aging_list()
                
                customer_code = str(receivable.pcode) +" - "+ str(receivable.pname)
                if customer_code not in datalist:
                    amount_list = self.aging_list()
                    datalist[customer_code] = {
                        'transactions': [],
                        'subtotal': amount_list,
                    }
                
                if balancecode == 'D':
                    amount = receivable.debitamount
                    
                    # days_past_due = calculate_aging(receivable.document_date, date_to_age)
                    months = self.diff_month(receivable.document_date, date_to_age)
                    
                    for bracket_name, condition in aging_brackets[:-1]:  # Skip "Over-Payments" for Credit transactions
                        if condition(months):
                            aging_parameter[bracket_name] = amount
                            datalist[customer_code]['subtotal']['amount_due'] += amount
                            datalist[customer_code]['subtotal'][bracket_name] += amount
                            
                            break
                else:
                    if receivable.document_refamount == 0:
                        amount = receivable.creditamount
                    else:
                        amount = receivable.document_refamount
                        print 'remaining balance', amount
                    
                    aging_parameter['overpayment'] = amount
                    datalist[customer_code]['subtotal']['amount_due'] += -abs(amount)
                    datalist[customer_code]['subtotal']['overpayment'] += -abs(amount)
                        
                aging_parameter['amount_due'] = amount
                
                details = {
                    'main_detail': receivable,
                    'aging_detail': aging_parameter
                }
                # print 'hoy', details['aging_detail']
                datalist[customer_code]['transactions'].append(details) 
                
            grandtotal = self.compute_grandtotal(datalist)
            
            data = datalist
        elif report == '1':
            for x, row in enumerate(data):
                if row.balancecode == 'D':
                    tdebit += row.debitamount
                else:
                    if row.document_refamount == 0:
                        tcredit += row.creditamount
                    else:
                        tcredit += abs(row.document_refamount)
        elif report == '3':
            for x, row in enumerate(data):
                if row.balancecode == 'D':
                    totalnetsales += row.debitamount
                else:
                    if row.document_refamount == 0:
                        totalnetsales += -abs(row.creditamount)
                    else:
                        totalnetsales += -abs(row.document_refamount)
                        
        context = {
            "title": title,
            "today": timezone.now(),
            "company": company,
            "data": data,
            "total_debit": tdebit,
            "total_credit": tcredit,
            "grandtotal": grandtotal,
            "totalnetsales": totalnetsales,
            "dfrom": dfrom,
            "dto": dto,
            "datefrom": datetime.datetime.strptime(dfrom, '%Y-%m-%d'),
            "dateto": datetime.datetime.strptime(dto, '%Y-%m-%d'),
            "username": request.user, 
        }
        
        if report == '1':
            return Render.render('nontrade/report/report_1.html', context)
        elif report == '2':
            return Render.render('nontrade/report/report_2.html', context)
        elif report == '3':
            return Render.render('nontrade/report/report_3.html', context)
        else:
            return Render.render('nontrade/report/report.html', context)
        

@method_decorator(login_required, name='dispatch')
class ManageARNonTradeView(ListView):
    model = Subledger
    template_name = 'nontrade/tagged_arnontrade_index.html'
    context_object_name = 'data_list'

    def dispatch(self, request, *args, **kwargs):
        return super(ListView, self).dispatch(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        context = super(ManageARNonTradeView, self).get_context_data(**kwargs)
        context['chart'] = Chartofaccount.objects.filter(isdeleted=0, accounttype='P', nontrade='Y').order_by('accountcode')
        
        return context


@csrf_exempt
def managearnontrade(request):
    viewhtml = ''
    context = {}
    try:
        dto = request.GET["dto"]
        dfrom = request.GET["dfrom"]
        transaction = 1
        chartofaccount = request.GET['chartofaccount']
        payeecode = request.GET['customercode']
        payeename = request.GET['customername']
        data = queryLedger(dto, dfrom, transaction, chartofaccount, payeecode, payeename, isNT=False, isClosedOnly=True)
        # queryTaggedARNonTrade(dto, dfrom)

        tdebit = 0
        tcredit = 0
        for val in data:
            tdebit += val.debitamount
            tcredit += val.creditamount
            
        context['data'] = data
        context['tdebit'] = tdebit
        context['tcredit'] = tcredit
        viewhtml = render_to_string('nontrade/tagged_arnontrade_transactions.html', context)
        
        data = {
            'status': 'success',
            'viewhtml': viewhtml
        }

        return JsonResponse(data)
    except Exception as e:
        print 'errpr', e
        return JsonResponse({
            'status': 'failed',
            'message': 'An error occured'
        })


@csrf_exempt
def untagarnontrade(request):
    try:
        id = request.POST.get('id')
        untag_data = Subledger.objects.filter(pk=id).first()

        amount = untag_data.amount

        tag_id = untag_data.tag_id
        main = Subledger.objects.filter(pk=tag_id).first()
        main_refamount = main.document_refamount
        main_amount = main.amount

        if Decimal(main_refamount) == 0.00:
            breakdown_count = Subledger.objects.filter(tag_id=tag_id).count()

            if breakdown_count > 2:
                main.document_refamount = amount
                main.save()
            else:

                main.document_reftype = None
                main.document_refnum = None
                main.document_refdate = None
                main.tag_id = None
                main.is_closed = 0
                main.save()

        else:
            new_refamount = Decimal(main_refamount) + Decimal(amount)

            if new_refamount == Decimal(main_amount):
                main.document_reftype = None
                main.document_refnum = None
                main.document_refdate = None
                main.document_refamount = 0.00
                main.tag_id = None
                main.is_closed = 0
            else:
                main.document_refamount = new_refamount

            main.save()
        
        untag_data.document_reftype = None
        untag_data.document_refnum = None
        untag_data.document_refdate = None
        untag_data.tag_id = None
        untag_data.is_closed = 0

        untag_data.save()

        response = {
            'status': 'success'
        }

    except Exception as e:
        response = {
            'status': 'failed',
            'message': "An unexpected error has occured: " + str(e)
        }
    
    return JsonResponse(response)