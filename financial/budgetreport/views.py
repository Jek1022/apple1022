from django.views.generic import View, TemplateView
from django.contrib.auth.decorators import login_required
from django.utils.decorators import method_decorator
from companyparameter.models import Companyparameter
from django.db import connection
from accountexpensebalance.models import Accountexpensebalance
from department.models import Department
from departmentbudget.models import Departmentbudget
from chartofaccount.models import Chartofaccount
from product.models import Product
from utils.mixins import ReportContentMixin
from easy_pdf.views import PDFTemplateView
from django.views.decorators.csrf import csrf_exempt
from django.db.models import Sum, Case, Value, When, F, Q
from financial.context_processors import namedtuplefetchall
import datetime
from financial.utils import Render
from django.utils import timezone
from django.template.loader import get_template
from django.http import HttpResponse, JsonResponse
import pandas as pd
from datetime import timedelta
import io
import xlsxwriter
import datetime
from django.template.loader import render_to_string

@method_decorator(login_required, name='dispatch')
class IndexView(TemplateView):
    model = Accountexpensebalance
    template_name = 'budgetreport/index.html'

    def get_context_data(self, **kwargs):
        context = super(TemplateView, self).get_context_data(**kwargs)

        context['product'] = Product.objects.filter(isdeleted=0).order_by('description')
        context['department'] = Department.objects.filter(isdeleted=0).order_by('code')
        context['expense'] = Chartofaccount.objects.filter(isdeleted=0, accounttype='T', pk__in=[527, 643, 745]).order_by('accountcode')

        return context

@method_decorator(login_required, name='dispatch')
class DeptBudgetInquiry(TemplateView):
    model = Accountexpensebalance
    template_name = 'budgetreport/inquiry/index.html'

    def get_context_data(self, **kwargs):
        context = super(TemplateView, self).get_context_data(**kwargs)
        context['chart'] = Chartofaccount.objects.filter(isdeleted=0, accounttype='P', main=5).order_by('accountcode')
        context['department'] = Department.objects.filter(isdeleted=0).order_by('code')

        return context

def transgenerate(request):
    dto = request.GET["dto"]
    dfrom = request.GET["dfrom"]
    chart = request.GET["chart"]
    department = request.GET["department"]

    context = {}

    print "transaction listing"

    ndto = datetime.datetime.strptime(dto, "%Y-%m-%d")
    todate = datetime.date(int(ndto.year), int(ndto.month), 10)
    toyear = todate.year
    tomonth = todate.month
    nfrom = datetime.datetime.strptime(dfrom, "%Y-%m-%d")
    fromdate = datetime.date(int(nfrom.year), int(nfrom.month), 10)
    fromyear = fromdate.year
    frommonth = fromdate.month

    budget = Departmentbudget.objects.filter(isdeleted=0,year=toyear)

    if chart != '':
        budget = budget.filter(chartofaccount_id=chart)
        print 'budget'

    if department != '':
        budget = budget.filter(department_id=department)
        print 'department'

    #budget = getBudget(tomonth, item)
    totalbudget = 0
    for item in budget:
        budgetdata = getBudget(tomonth, item)
        totalbudget += budgetdata[1]

    data = query_transaction(dto, dfrom, chart, department)

    debit = 0
    credit = 0
    total = 0
    totalvariance = 0
    for item in data:
        debit += item.debitamount
        credit += item.creditamount

    total = float(debit) - float(credit)
    totalvariance = float(totalbudget) - float(total)

    context['result'] = data
    context['dto'] = dto
    context['dfrom'] = dfrom
    context['debit'] = debit
    context['credit'] = credit
    context['total'] = total
    viewhtml = render_to_string('budgetreport/inquiry/transaction_result.html', context)


    data = {
        'status': 'success',
        'viewhtml': viewhtml,
        'totalbudget': totalbudget,
        'totalactual': total,
        'totalvariance': totalvariance,
    }
    return JsonResponse(data)

# @method_decorator(login_required, name='dispatch')
# class GeneratePDF(View):
#     def get(self, request):
#         company = Companyparameter.objects.all().first()
#         list = []
#         total = []
#         report = request.GET['report']
#         filter = request.GET['filter']
#         type = request.GET['type']
#         dfrom = request.GET['from']
#         dto = request.GET['to']
#         department = request.GET['department']
#         product = request.GET['product']
#         title = "Budget Monitoring Report"
#         subtitle = ""
#         typetext = "Summary"
#         filtertext = "Cost of Sales"
#
#         ndto = datetime.datetime.strptime(dto, "%Y-%m-%d")
#         todate = datetime.date(int(ndto.year), int(ndto.month), 10)
#         toyear = todate.year
#         tomonth = todate.month
#         nfrom = datetime.datetime.strptime(dfrom, "%Y-%m-%d")
#         fromdate = datetime.date(int(nfrom.year), int(nfrom.month), 10)
#         fromyear = fromdate.year
#         frommonth = fromdate.month
#
#         accountcode = 0
#         curmon_var = 0
#         curmon_var_per = 0
#         yearcur_var = 0
#         yearcur_var_per = 0
#         yearprev_var = 0
#         yearprev_var_per = 0
#         subtotal_curmon_bud = 0
#         subtotal_curmon_act = 0
#         subtotal_curmon_var = 0
#         subtotal_curmon_var_per = 0
#         subtotal_curyear_bud = 0
#         subtotal_curyear_act = 0
#         subtotal_curyear_var = 0
#         subtotal_curyear_var_per = 0
#         subtotal_prevyear_act = 0
#         subtotal_prevyear_var = 0
#         subtotal_prevyear_var_per = 0
#         total_curmon_bud = 0
#         total_curmon_act = 0
#         total_curmon_var = 0
#         total_curmon_var_per = 0
#         total_curyear_bud = 0
#         total_curyear_act = 0
#         total_curyear_var = 0
#         total_curyear_var_per = 0
#         total_prevyear_act = 0
#         total_prevyear_var = 0
#         total_prevyear_var_per = 0
#
#         counter = 0
#
#         list = Accountexpensebalance.objects.filter(isdeleted=0)[:0]
#
#         if filter == '1':
#             filtertext = "Cost of Sales"
#         elif filter == '2':
#             filtertext = "General & Administrative"
#         elif filter == '3':
#             filtertext = "Selling"
#         else:
#             filtertext = "All"
#
#         if type == '1':
#             typetext = "Summary"
#         else:
#             typetext = "Detailed"
#
#         new_list = []
#
#         if report == '1':
#             subtitle = "Budget Status By Department/Section ( " + filtertext + " - " + typetext + " )"
#             data = query_bugdet_status_by_department(filter, type, fromyear, frommonth, toyear, tomonth, department, product)
#
#             df = pd.DataFrame(data)
#
#             if type == '2':
#                 if data:
#                     for dept, department in df.fillna('NaN').sort_values(by=['deptcode', 'chartgroupaccountcode', 'chartsubgroupaccountcode', 'accountcode'], ascending=True).groupby(['deptcode', 'departmentname']):
#                         total_curmon_bud = 0
#                         total_curmon_act = 0
#                         total_curmon_var = 0
#                         total_curyear_bud = 0
#                         total_curyear_act = 0
#                         total_curyear_var = 0
#                         total_prevyear_act = 0
#                         total_prevyear_var = 0
#                         total_prevyear_var_per = 0
#                         for group, chartgroup in department.fillna('NaN').sort_values(by=['deptcode', 'chartgroupaccountcode', 'chartsubgroupaccountcode', 'accountcode'], ascending=True).groupby(['chartgroup']):
#                             counter += 1
#                             accountcode = 0
#                             subtotal_curmon_bud = 0
#                             subtotal_curmon_act = 0
#                             subtotal_curmon_var = 0
#                             subtotal_curyear_bud = 0
#                             subtotal_curyear_act = 0
#                             subtotal_curyear_var = 0
#                             subtotal_prevyear_act = 0
#                             subtotal_prevyear_var = 0
#                             subtotal_prevyear_var_per = 0
#                             for subgroup, chartsubgroup in chartgroup.sort_values(by=['accountcode'], ascending=True).groupby(['chartsubgroup']):
#
#                                 for data, item in chartsubgroup.iterrows():
#
#                                     budget = getBudget(tomonth, item)
#                                     curmon_var = float(item.actualcurmonamount) - float(budget[0])
#                                     if float(budget[0]):
#                                         curmon_var_per = (float(curmon_var) / float(budget[0])) * 100
#                                     else:
#                                         curmon_var_per = 0
#
#                                     yearcur_var = float(item.actualcuryearamount) - float(budget[1])
#                                     if float(budget[1]):
#                                         yearcur_var_per = (float(yearcur_var) / float(budget[1]))* 100
#                                     else:
#                                         yearcur_var_per = 0
#
#                                     yearprev_var = float(item.actualcuryearamount) - float(item.actualprevyearamount)
#                                     if float(item.actualprevyearamount):
#                                         yearprev_var_per = (float(yearprev_var) / float(item.actualprevyearamount)) * 100
#                                     else:
#                                         yearprev_var_per = 0
#
#                                     new_list.append({'chartgroup': group, 'chartsubgroup': subgroup, 'accountcode': item.accountcode,
#                                                      'chartofaccount': item.description, 'deptcode': item.deptcode, 'department': item.departmentname,
#                                                      'curmon_bud': budget[0], 'curmon_act': item.actualcurmonamount, 'curmon_var': curmon_var, 'curmon_var_per': curmon_var_per,
#                                                      'yearcur_bud': budget[1], 'yearcur_act': item.actualcuryearamount, 'yearcur_var': yearcur_var, 'yearcur_var_per': yearcur_var_per,
#                                                      'yearprev_act': item.actualprevyearamount, 'yearprev_var': yearprev_var, 'yearprev_var_per': yearprev_var_per, 'counter': counter})
#
#                                     subtotal_curmon_bud += budget[0]
#                                     subtotal_curmon_act += item.actualcurmonamount
#                                     subtotal_curmon_var += curmon_var
#                                     subtotal_curyear_bud +=  budget[1]
#                                     subtotal_curyear_act += item.actualcuryearamount
#                                     subtotal_curyear_var += yearcur_var
#                                     subtotal_prevyear_act += item.actualprevyearamount
#                                     subtotal_prevyear_var += yearprev_var
#
#                                     total_curmon_bud += budget[0]
#                                     total_curmon_act += item.actualcurmonamount
#                                     total_curmon_var += curmon_var
#                                     total_curyear_bud += budget[1]
#                                     total_curyear_act += item.actualcuryearamount
#                                     total_curyear_var += yearcur_var
#                                     total_prevyear_act += item.actualprevyearamount
#                                     total_prevyear_var += yearprev_var
#
#                             if float(subtotal_curmon_bud):
#                                 subtotal_curmon_var_per = (float(subtotal_curmon_var) / float(subtotal_curmon_bud)) * 100
#                             if float(subtotal_curyear_bud):
#                                 subtotal_curyear_var_per = (float(subtotal_curyear_var) / float(subtotal_curyear_bud))* 100
#                             if float(subtotal_prevyear_act):
#                                 subtotal_prevyear_var_per = (float(subtotal_prevyear_var) / float(subtotal_prevyear_act)) * 100
#
#                             new_list.append({'chartgroup': group, 'chartsubgroup': 'subtotal', 'accountcode': accountcode,
#                                              'chartofaccount': 'subtotal', 'deptcode': dept[0], 'department': dept[1],
#                                              'curmon_bud': subtotal_curmon_bud, 'curmon_act': subtotal_curmon_act, 'curmon_var': subtotal_curmon_var, 'curmon_var_per': subtotal_curmon_var_per,
#                                              'yearcur_bud': subtotal_curyear_bud, 'yearcur_act': subtotal_curyear_act, 'yearcur_var': subtotal_curyear_var, 'yearcur_var_per': subtotal_curyear_var_per,
#                                              'yearprev_act': subtotal_prevyear_act, 'yearprev_var': subtotal_prevyear_var, 'yearprev_var_per': subtotal_prevyear_var_per,
#                                              'counter': counter + 1})
#
#                         if float(total_curmon_bud):
#                             total_curmon_var_per = (float(total_curmon_var) / float(total_curmon_bud)) * 100
#                         if float(total_curyear_bud):
#                             total_curyear_var_per = (float(total_curyear_var) / float(total_curyear_bud)) * 100
#                         if float(total_prevyear_act):
#                             total_prevyear_var_per = (float(total_prevyear_var) / float(total_prevyear_act)) * 100
#
#                         new_list.append({'chartgroup': 'total', 'chartsubgroup': 'total', 'accountcode': accountcode,
#                                          'chartofaccount': 'total', 'deptcode': dept[0], 'department': dept[1],
#                                          'curmon_bud': total_curmon_bud, 'curmon_act': total_curmon_act, 'curmon_var': total_curmon_var, 'curmon_var_per': total_curmon_var_per,
#                                          'yearcur_bud': total_curyear_bud, 'yearcur_act': total_curyear_act, 'yearcur_var': total_curyear_var, 'yearcur_var_per': total_curyear_var_per,
#                                          'yearprev_act': total_prevyear_act, 'yearprev_var': total_prevyear_var, 'yearprev_var_per': total_prevyear_var_per,
#                                          'counter': counter + 1})
#             else:
#                 if data:
#                     for dept, department in df.fillna('NaN').sort_values(by=['deptcode', 'chartgroupaccountcode', 'chartsubgroupaccountcode', 'accountcode'], ascending=True).groupby(['deptcode', 'departmentname']):
#                         total_curmon_bud = 0
#                         total_curmon_act = 0
#                         total_curmon_var = 0
#                         total_curyear_bud = 0
#                         total_curyear_act = 0
#                         total_curyear_var = 0
#                         total_prevyear_act = 0
#                         total_prevyear_var = 0
#                         total_prevyear_var_per = 0
#                         for group, chartgroup in department.fillna('NaN').sort_values(by=['accountcode'],ascending=True).groupby(['chartgroup']):
#                             counter += 1
#                             accountcode = 0
#                             subtotal_curmon_bud = 0
#                             subtotal_curmon_act = 0
#                             subtotal_curmon_var = 0
#                             subtotal_curyear_bud = 0
#                             subtotal_curyear_act = 0
#                             subtotal_curyear_var = 0
#                             subtotal_prevyear_act = 0
#                             subtotal_prevyear_var = 0
#                             subtotal_prevyear_var_per = 0
#                             for data, item in chartgroup.iterrows():
#                                 counter += 1
#                                 budget = getBudget(tomonth, item)
#
#                                 curmon_var = float(item.actualcurmonamount) - float(budget[0])
#                                 if float(budget[0]):
#                                     curmon_var_per = (float(curmon_var) / float(budget[0])) * 100
#                                 else:
#                                     curmon_var_per = 0
#
#                                 yearcur_var = float(item.actualcuryearamount) - float(budget[1])
#                                 if float(budget[1]):
#                                     yearcur_var_per = (float(yearcur_var) / float(budget[1])) * 100
#                                 else:
#                                     yearcur_var_per = 0
#
#                                 yearprev_var = float(item.actualcuryearamount) - float(item.actualprevyearamount)
#                                 if float(item.actualprevyearamount):
#                                     yearprev_var_per = (float(yearprev_var) / float(item.actualprevyearamount)) * 100
#                                 else:
#                                     yearprev_var_per = 0
#
#                                 new_list.append({'chartgroup': group, 'chartsubgroup': item.chartsubgroup, 'accountcode': item.accountcode,'chartofaccount': item.description,
#                                                  'deptcode': item.deptcode, 'department': item.departmentname,
#                                                  'curmon_bud': budget[0], 'curmon_act': item.actualcurmonamount, 'curmon_var': curmon_var, 'curmon_var_per': curmon_var_per,
#                                                  'yearcur_bud': budget[1], 'yearcur_act': item.actualcuryearamount, 'yearcur_var': yearcur_var, 'yearcur_var_per': yearcur_var_per,
#                                                  'yearprev_act': item.actualprevyearamount, 'yearprev_var': yearprev_var, 'yearprev_var_per': yearprev_var_per, 'counter': counter})
#
#                                 subtotal_curmon_bud += budget[0]
#                                 subtotal_curmon_act += item.actualcurmonamount
#                                 subtotal_curmon_var += curmon_var
#                                 subtotal_curyear_bud += budget[1]
#                                 subtotal_curyear_act += item.actualcuryearamount
#                                 subtotal_curyear_var += yearcur_var
#                                 subtotal_prevyear_act += item.actualprevyearamount
#                                 subtotal_prevyear_var += yearprev_var
#
#                                 total_curmon_bud += budget[0]
#                                 total_curmon_act += item.actualcurmonamount
#                                 total_curmon_var += curmon_var
#                                 total_curyear_bud += budget[1]
#                                 total_curyear_act += item.actualcuryearamount
#                                 total_curyear_var += yearcur_var
#                                 total_prevyear_act += item.actualprevyearamount
#                                 total_prevyear_var += yearprev_var
#
#                             if float(subtotal_curmon_bud):
#                                 subtotal_curmon_var_per = (float(subtotal_curmon_var) / float(
#                                     subtotal_curmon_bud)) * 100
#                             if float(subtotal_curyear_bud):
#                                 subtotal_curyear_var_per = (float(subtotal_curyear_var) / float(
#                                     subtotal_curyear_bud)) * 100
#                             if float(subtotal_prevyear_act):
#                                 subtotal_prevyear_var_per = (float(subtotal_prevyear_var) / float(
#                                     subtotal_prevyear_act)) * 100
#
#                             new_list.append({'chartgroup': group, 'chartsubgroup': 'subtotal', 'accountcode': accountcode,
#                                              'chartofaccount': 'subtotal', 'deptcode': dept[0], 'department': dept[1],
#                                              'curmon_bud': subtotal_curmon_bud, 'curmon_act': subtotal_curmon_act,
#                                              'curmon_var': subtotal_curmon_var, 'curmon_var_per': subtotal_curmon_var_per,
#                                              'yearcur_bud': subtotal_curyear_bud, 'yearcur_act': subtotal_curyear_act,
#                                              'yearcur_var': subtotal_curyear_var, 'yearcur_var_per': subtotal_curyear_var_per,
#                                              'yearprev_act': subtotal_prevyear_act, 'yearprev_var': subtotal_prevyear_var,
#                                              'yearprev_var_per': subtotal_prevyear_var_per,
#                                              'counter': counter + 1})
#
#                         if float(total_curmon_bud):
#                             total_curmon_var_per = (float(total_curmon_var) / float(total_curmon_bud)) * 100
#                         if float(total_curyear_bud):
#                             total_curyear_var_per = (float(total_curyear_var) / float(total_curyear_bud)) * 100
#                         if float(total_prevyear_act):
#                             total_prevyear_var_per = (float(total_prevyear_var) / float(total_prevyear_act)) * 100
#
#                         new_list.append({'chartgroup': 'total', 'chartsubgroup': 'total', 'accountcode': accountcode,
#                                          'chartofaccount': 'total', 'deptcode': dept[0], 'department': dept[1],
#                                          'curmon_bud': total_curmon_bud, 'curmon_act': total_curmon_act, 'curmon_var': total_curmon_var, 'curmon_var_per': total_curmon_var_per,
#                                          'yearcur_bud': total_curyear_bud, 'yearcur_act': total_curyear_act, 'yearcur_var': total_curyear_var, 'yearcur_var_per': total_curyear_var_per,
#                                          'yearprev_act': total_prevyear_act, 'yearprev_var': total_prevyear_var, 'yearprev_var_per': total_prevyear_var_per,
#                                          'counter': counter + 1})
#         print report
#         context = {
#             "title": title,
#             "subtitle": subtitle,
#             "today": timezone.now(),
#             "company": company,
#             "list": new_list,
#             "filtertext": filtertext,
#             "typetext": typetext,
#             "total": total,
#             "username": request.user,
#         }
#
#         if report == '1':
#             if type == '2':
#                 return Render.render('budgetreport/report_2_d.html', context)
#             else:
#                 print 'dito'
#                 return Render.render('budgetreport/department_summary.html', context)
#         else:
#             return Render.render('budgetreport/report_1.html', context)

def getBudget(tomonth, item):
    if tomonth == 1:
        monbudget = float(item.mjan)
        yearbudget = float(item.mjan)
    elif tomonth == 2:
        monbudget = float(item.mfeb)
        yearbudget = float(item.mjan) + float(item.mfeb)
    elif tomonth == 3:
        monbudget = float(item.mmar)
        yearbudget = float(item.mjan) + float(item.mfeb) + float(item.mmar)
    elif tomonth == 4:
        monbudget = float(item.mapr)
        yearbudget = float(item.mjan) + float(item.mfeb) + float(item.mmar) + float(item.mapr)
    elif tomonth == 5:
        monbudget = float(item.mmay)
        yearbudget = float(item.mjan) + float(item.mfeb) + float(item.mmar) + float(item.mapr) + float(item.mmay)
    elif tomonth == 6:
        monbudget = float(item.mjun)
        yearbudget = float(item.mjan) + float(item.mfeb) + float(item.mmar) + float(item.mapr) + float(item.mmay) \
                     + float(item.mjun)
    elif tomonth == 7:
        monbudget = float(item.mjul)
        yearbudget = float(item.mjan) + float(item.mfeb) + float(item.mmar) + float(item.mapr) + float(item.mmay) \
                     + float(item.mjun) + float(item.mjul)
    elif tomonth == 8:
        monbudget = float(item.maug)
        yearbudget = float(item.mjan) + float(item.mfeb) + float(item.mmar) + float(item.mapr) + float(item.mmay) \
                     + float(item.mjun) + float(item.mjul) + float(item.maug)
    elif tomonth == 9:
        monbudget = float(item.msep)
        yearbudget = float(item.mjan) + float(item.mfeb) + float(item.mmar) + float(item.mapr) + float(item.mmay) \
                     + float(item.mjun) + float(item.mjul) + float(item.maug) + float(item.msep)
    elif tomonth == 10:
        monbudget = float(item.moct)
        yearbudget = float(item.mjan) + float(item.mfeb) + float(item.mmar) + float(item.mapr) + float(item.mmay) \
                     + float(item.mjun) + float(item.mjul) + float(item.maug) + float(item.msep) + float(item.moct)
    elif tomonth == 11:
        monbudget = float(item.mnov)
        yearbudget = float(item.mjan) + float(item.mfeb) + float(item.mmar) + float(item.mapr) + float(item.mmay) \
                     + float(item.mjun) + float(item.mjul) + float(item.maug) + float(item.msep) + float(item.moct) + float(item.mnov)
    elif tomonth == 12:
        monbudget = float(item.mdec)
        yearbudget = float(item.mjan) + float(item.mfeb) + float(item.mmar) + float(item.mapr) + float(item.mmay) \
                     + float(item.mjun) + float(item.mjul) + float(item.maug) + float(item.msep) + float(item.moct) + float(item.mnov) + float(item.mdec)
    else:
        monbudget = float(item.mjan)
        yearbudget = float(item.mjan)

    return [monbudget, yearbudget]

def query_bugdet_status_by_department(filter, type, fromyear, frommonth, toyear, tomonth, department, product):
    prevyear = int(toyear) - 1
    condepartment = ""
    condepartment2 = ""
    conbudgroup = ""
    conproduc = ""
    contype = ""
    confilter = ""

    if filter == '1':
        confilter = "AND chart.main = 5 AND chart.clas = 1"
    elif filter == '2':
        confilter = "AND chart.main = 5 AND chart.clas = 2"
    else:
        confilter = "AND chart.main = 5 AND chart.clas = 3"

    if department:
        #condepartment = "AND deptbud.department_id IN (48,22)"
        #condepartment2 = "AND acctbal.department_id IN (48,22)"
        condepartment = "AND deptbud.department_id = "+str(department)+" "
        condepartment2 = "AND acctbal.department_id = "+str(department)+" "
    if type == '2':
        print 'detailed'
        contype = "GROUP BY z.chartgroup, z.chartsubgroup, z.accountcode, z.deptcode"
        conbudgroup = "GROUP BY chartgroup.title, chartsubgroup.title, chart.accountcode, dept.code"
    elif type == '1':
        print 'summary'
        contype = "GROUP BY z.chartgroup, z.chartsubgroup, z.deptcode"
        conbudgroup = "GROUP BY chartgroup.title, chartsubgroup.title, dept.code"

    print "raw query"
    ''' Create query '''
    cursor = connection.cursor()

    query = "SELECT z.accountcode, z.chartgroup, z.chartgroupaccountcode, z.chartsubgroup, z.chartsubgroupaccountcode, " \
            "z.main, z.clas, z.item, z.cont, z.sub, z.description, " \
            "z.deptcode, z.departmentname, z.year, " \
            "z.mjan, z.mfeb, z.mmar, z.mapr,  z.mmay,  z.mjun, " \
            "z.mjul, z.maug, z.msep, z.moct, z.mnov, z.mdec, " \
            "SUM(IFNULL(z.actualcurmonamount, 0)) AS actualcurmonamount, SUM(IFNULL(z.actualcuryearamount, 0)) AS actualcuryearamount, SUM(IFNULL(z.actualprevyearamount, 0)) AS actualprevyearamount " \
            "FROM ( " \
            "SELECT chartgroup.title AS chartgroup, chartgroup.accountcode AS chartgroupaccountcode, chartsubgroup.title AS chartsubgroup, chartsubgroup.accountcode AS chartsubgroupaccountcode, " \
            "chart.main, chart.clas, chart.item, chart.cont, chart.sub, chart.accountcode, chart.description, " \
            "dept.code AS deptcode, dept.departmentname, deptbud.year, " \
            "SUM(deptbud.mjan) AS mjan, SUM(deptbud.mfeb) AS mfeb, SUM(deptbud.mmar) AS mmar, SUM(deptbud.mapr) AS mapr,  SUM(deptbud.mmay) AS mmay,  SUM(deptbud.mjun) AS mjun, " \
            "SUM(deptbud.mjul) AS mjul, SUM(deptbud.maug) AS maug, SUM(deptbud.msep) AS msep, SUM(deptbud.moct) AS moct, SUM(deptbud.mnov) AS mnov, SUM(deptbud.mdec) AS mdec, " \
            "actualcurmon.year AS actualcurmonyear, actualcurmon.month AS actualcurmonmonth, actualcurmon.amount AS actualcurmonamount, " \
            "'' AS actualcuryear, 0 AS actualcuryearamount, " \
            "'' AS actualprevyear, 0 AS actualprevyearamount " \
            "FROM chartofaccount AS chart " \
            "LEFT OUTER JOIN departmentbudget AS deptbud ON deptbud.chartofaccount_id = chart.id " \
            "LEFT OUTER JOIN department AS dept ON dept.id = deptbud.department_id " \
            "LEFT OUTER JOIN chartofaccount AS chartgroup ON (chartgroup.main = chart.main AND chartgroup.clas = chart.clas " \
            "AND chartgroup.item = chart.item AND chartgroup.cont = 0 " \
            "AND chartgroup.sub = 000000 AND chartgroup.accounttype = 'T') " \
            "LEFT OUTER JOIN chartofaccount AS chartsubgroup ON (chartsubgroup.main = chart.main AND chartsubgroup.clas = chart.clas " \
            "AND chartsubgroup.item = chart.item AND chartsubgroup.cont = chart.cont " \
            "AND chartsubgroup.sub = RPAD(SUBSTR(chart.sub, 1, 1), 6, 0)) " \
            "LEFT OUTER JOIN ( " \
            "SELECT a.year, a.month, SUM(IF (a.code = 'C', (a.amount * -1), a.amount)) AS amount, a.code, a.chartofaccount_id, a.department_id " \
            "FROM accountexpensebalance AS a " \
            "WHERE a.year = "+str(toyear)+" AND a.month = "+str(tomonth)+" " \
            "GROUP BY a.chartofaccount_id, a.department_id " \
            ") AS actualcurmon ON (actualcurmon.chartofaccount_id = deptbud.chartofaccount_id AND actualcurmon.department_id = deptbud.department_id) " \
            "WHERE deptbud.isdeleted = 0 "+str(condepartment)+" " \
            " "+str(confilter)+" " \
            " "+str(conbudgroup)+" " \
            "UNION " \
            "SELECT chartgroup.title AS chartgroup, chartgroup.accountcode AS chartgroupaccountcode, chartsubgroup.title AS chartsubgroup, chartsubgroup.accountcode AS chartsubgroupaccountcode, " \
            "chart.main, chart.clas, chart.item, chart.cont, chart.sub, chart.accountcode, chart.description, " \
            "dept.code AS deptcode, dept.departmentname, acctbal.year, " \
            "0 AS mjan, 0 AS mfeb, 0 AS mmar, 0 AS mapr, 0 AS mmay,  0 AS mjun, " \
            "0 AS mjul, 0 AS maug, 0 AS msep, 0 AS moct, 0 AS mnov, 0 AS mdec, " \
            "acctbal.year AS actualcurmonyear, "+str(tomonth)+" AS actualcurmonmonth, IFNULL(actualcurmon.amount, 0) AS actualcurmonamount, " \
            "acctbal.year AS actualcuryear, SUM(IF (acctbal.code = 'C', (acctbal.amount * -1), acctbal.amount)) AS actualcuryearamount, " \
            "'' AS actualprevyear, 0 AS actualprevyearamount " \
            "FROM chartofaccount AS chart " \
            "LEFT OUTER JOIN accountexpensebalance AS acctbal ON acctbal.chartofaccount_id = chart.id " \
            "LEFT OUTER JOIN department AS dept ON dept.id = acctbal.department_id " \
            "LEFT OUTER JOIN chartofaccount AS chartgroup ON (chartgroup.main = chart.main AND chartgroup.clas = chart.clas " \
            "AND chartgroup.item = chart.item AND chartgroup.cont = 0 " \
            "AND chartgroup.sub = 000000 AND chartgroup.accounttype = 'T') " \
            "LEFT OUTER JOIN chartofaccount AS chartsubgroup ON (chartsubgroup.main = chart.main AND chartsubgroup.clas = chart.clas " \
            "AND chartsubgroup.item = chart.item AND chartsubgroup.cont = chart.cont " \
            "AND chartsubgroup.sub = RPAD(SUBSTR(chart.sub, 1, 1), 6, 0)) " \
            "LEFT OUTER JOIN (" \
            "SELECT acctbal.id, acctbal.year, acctbal.month, (IF (acctbal.code = 'C', (acctbal.amount * -1), acctbal.amount)) AS amount, acctbal.code, acctbal.chartofaccount_id, acctbal.department_id " \
            "FROM accountexpensebalance AS acctbal " \
            "WHERE acctbal.year = "+str(fromyear)+" AND acctbal.month = "+str(frommonth)+" "+str(condepartment2)+" " \
            "AND acctbal.chartofaccount_id NOT IN(SELECT DISTINCT(chart.id) FROM chartofaccount AS chart " \
            "INNER JOIN departmentbudget AS deptbud ON deptbud.chartofaccount_id = chart.id WHERE acctbal.isdeleted = 0 "+str(condepartment)+") " \
            "GROUP BY acctbal.chartofaccount_id, acctbal.department_id) AS actualcurmon ON (actualcurmon.chartofaccount_id = acctbal.chartofaccount_id AND actualcurmon.department_id = acctbal.department_id)	 " \
            "WHERE acctbal.year >= "+str(fromyear)+" AND acctbal.year <= "+str(toyear)+" " \
            "AND acctbal.month >= "+str(frommonth)+" AND acctbal.month <= "+str(tomonth)+" " \
            " "+str(confilter)+" " \
            " "+str(condepartment2)+" " \
            "UNION " \
            "SELECT chartgroup.title AS chartgroup, chartgroup.accountcode AS chartgroupaccountcode, chartsubgroup.title AS chartsubgroup, chartsubgroup.accountcode AS chartsubgroupaccountcode, " \
            "chart.main, chart.clas, chart.item, chart.cont, chart.sub, chart.accountcode, chart.description, " \
            "dept.code AS deptcode, dept.departmentname, acctbal.year, " \
            "0 AS mjan, 0 AS mfeb, 0 AS mmar, 0 AS mapr, 0 AS mmay,  0 AS mjun, " \
            "0 AS mjul, 0 AS maug, 0 AS msep, 0 AS moct, 0 AS mnov, 0 AS mdec, " \
            "'' AS actualcurmonyear, acctbal.month AS actualcurmonmonth, 0 AS actualcurmonamount, " \
            "'', 0 AS actualcuryearamount, " \
            "acctbal.year AS actualprevyear, SUM(IF (acctbal.code = 'C', (acctbal.amount * -1), acctbal.amount)) AS actualprevyearamount " \
            "FROM chartofaccount AS chart " \
            "LEFT OUTER JOIN accountexpensebalance AS acctbal ON acctbal.chartofaccount_id = chart.id " \
            "LEFT OUTER JOIN department AS dept ON dept.id = acctbal.department_id " \
            "LEFT OUTER JOIN chartofaccount AS chartgroup ON (chartgroup.main = chart.main AND chartgroup.clas = chart.clas " \
            "AND chartgroup.item = chart.item AND chartgroup.cont = 0 " \
            "AND chartgroup.sub = 000000 AND chartgroup.accounttype = 'T') " \
            "LEFT OUTER JOIN chartofaccount AS chartsubgroup ON (chartsubgroup.main = chart.main AND chartsubgroup.clas = chart.clas " \
            "AND chartsubgroup.item = chart.item AND chartsubgroup.cont = chart.cont " \
            "AND chartsubgroup.sub = RPAD(SUBSTR(chart.sub, 1, 1), 6, 0)) " \
            "WHERE acctbal.year >= "+str(prevyear)+" AND acctbal.year <= "+str(prevyear)+" " \
            "AND acctbal.month >= "+str(frommonth)+" AND acctbal.month <= "+str(tomonth)+" " \
            " "+str(confilter)+" " \
            " "+str(condepartment2)+" " \
            "GROUP BY chartgroup.title, chartsubgroup.title, chart.accountcode, dept.code) AS z " \
            "WHERE z.chartgroup IS NOT NULL " \
            " "+str(contype)+" " \
            "ORDER BY z.deptcode, z.accountcode , z.chartgroup, z.chartsubgroup"

    #print query

    cursor.execute(query)
    result = namedtuplefetchall(cursor)

    return result


def query_transaction(dto, dfrom, chart, department):
    print "Transaction Query Department Budget"
    ''' Create query '''
    cursor = connection.cursor()

    chart_condition = ''
    department_condition = ''

    if chart != '':
        chart_condition = "AND d.chartofaccount_id = '" + str(chart) + "'"
    if department != '':
        department_condition = "AND d.department_id = '" + str(department) + "'"

    query = "SELECT z.tran, z.item_counter, z.ap_num AS tnum, z.ap_date AS tdate, z.debitamount, z.creditamount, z.balancecode, z.apstatus AS transtatus, z.status AS status, " \
            "z.particulars, bank.code AS bank, chart.accountcode, chart.description AS chartofaccount, cust.code AS custcode, cust.name AS customer, dept.code AS deptcode, dept.departmentname AS department, " \
            "emp.code AS empcode, CONCAT(IFNULL(emp.firstname, ''), ' ', IFNULL(emp.lastname, '')) AS employee, inpvat.code AS inpvatcode, inpvat.description AS inputvat, " \
            "outvat.code AS outvatcode, outvat.description AS outputvat, prod.code AS prodcode, prod.description AS product, " \
            "supp.code AS suppcode, supp.name AS supplier, vat.code AS vatcode, vat.description AS vat, wtax.code AS wtaxcode, wtax.description AS wtax " \
            "FROM ( " \
            "SELECT 'AP' AS tran, d.item_counter, d.ap_num, d.ap_date, d.debitamount, d.creditamount, d.balancecode, d.ataxcode_id, m.particulars, " \
            "d.bankaccount_id, d.branch_id, d.chartofaccount_id, d.customer_id, d.department_id, d.employee_id, d.inputvat_id, " \
            "d.outputvat_id, d.product_id, d.supplier_id, d.vat_id, d.wtax_id, m.apstatus, m.status " \
            "FROM apdetail AS d " \
            "LEFT OUTER JOIN apmain AS m ON m.id = d.apmain_id " \
            "WHERE DATE(d.ap_date) >= '"+str(dfrom)+"' AND DATE(d.ap_date) <= '"+str(dto)+"' AND m.apstatus = 'R' AND m.status = 'O' " \
            +str(chart_condition)+" "+str(department_condition)+"" \
            "UNION " \
            "SELECT 'CV' AS tran, d.item_counter, d.cv_num, d.cv_date, d.debitamount, d.creditamount, d.balancecode, d.ataxcode_id, m.particulars, " \
            "d.bankaccount_id, d.branch_id, d.chartofaccount_id, d.customer_id, d.department_id, d.employee_id, d.inputvat_id, " \
            "d.outputvat_id, d.product_id, d.supplier_id, d.vat_id, d.wtax_id, m.cvstatus, m.status	" \
            "FROM cvdetail AS d " \
            "LEFT OUTER JOIN cvmain AS m ON m.id = d.cvmain_id " \
            "WHERE DATE(d.cv_date) >= '"+str(dfrom)+"' AND DATE(d.cv_date) <= '"+str(dto)+"' AND m.cvstatus = 'R' AND m.status = 'O' " \
            +str(chart_condition)+" "+str(department_condition)+"" \
            "UNION " \
            "SELECT 'JV' AS tran, d.item_counter, d.jv_num, d.jv_date, d.debitamount, d.creditamount, d.balancecode, d.ataxcode_id, m.particular," \
            "d.bankaccount_id, d.branch_id, d.chartofaccount_id, d.customer_id, d.department_id, d.employee_id, d.inputvat_id, " \
            "d.outputvat_id, d.product_id, d.supplier_id, d.vat_id, d.wtax_id, m.jvstatus, m.status " \
            "FROM jvdetail AS d " \
            "LEFT OUTER JOIN jvmain AS m ON m.id = d.jvmain_id " \
            "WHERE DATE(d.jv_date) >= '"+str(dfrom)+"' AND DATE(d.jv_date) <= '"+str(dto)+"' AND m.jvstatus = 'R' AND m.status = 'O' " \
            +str(chart_condition)+" "+str(department_condition)+"" \
            "UNION " \
            "SELECT 'OR' AS tran, d.item_counter, m.ornum, m.ordate, d.debitamount, d.creditamount, d.balancecode, d.ataxcode_id, m.particulars, " \
            "d.bankaccount_id, d.branch_id, d.chartofaccount_id, d.customer_id, d.department_id, d.employee_id, d.inputvat_id, " \
            "d.outputvat_id, d.product_id, d.supplier_id, d.vat_id, d.wtax_id, m.orstatus, m.status " \
            "FROM ormain AS m " \
            "LEFT OUTER JOIN ordetail AS d ON m.id = d.ormain_id " \
            "WHERE DATE(m.ordate) >= '"+str(dfrom)+"' AND DATE(m.ordate) <= '"+str(dto)+"' AND m.orstatus = 'R' AND m.status = 'O' " \
            +str(chart_condition)+" "+str(department_condition)+") AS z " \
            "LEFT OUTER JOIN bankaccount AS bank ON bank.id = z.bankaccount_id " \
            "LEFT OUTER JOIN chartofaccount AS chart ON chart.id = z.chartofaccount_id " \
            "LEFT OUTER JOIN customer AS cust ON cust.id = z.customer_id " \
            "LEFT OUTER JOIN department AS dept ON dept.id = z.department_id " \
            "LEFT OUTER JOIN employee AS emp ON emp.id = z.employee_id " \
            "LEFT OUTER JOIN inputvat AS inpvat ON inpvat.id = z.inputvat_id " \
            "LEFT OUTER JOIN outputvat AS outvat ON outvat.id = z.outputvat_id " \
            "LEFT OUTER JOIN product AS prod ON prod.id = z.product_id " \
            "LEFT OUTER JOIN supplier AS supp ON supp.id = z.supplier_id " \
            "LEFT OUTER JOIN vat AS vat ON vat.id = z.vat_id " \
            "LEFT OUTER JOIN wtax AS wtax ON wtax.id = z.wtax_id WHERE chart.main = 5 " \
            "ORDER BY z.tran, z.ap_num, z.ap_date, z.item_counter"

    cursor.execute(query)
    result = namedtuplefetchall(cursor)

    return result


@method_decorator(login_required, name='dispatch')
class GenerateTransExcel(View):
    def get(self, request):
        company = Companyparameter.objects.all().first()
        chartofaccount = []
        dept = []
        dto = request.GET["dto"]
        dfrom = request.GET["dfrom"]
        chart = request.GET["chart"]
        department = request.GET["department"]

        context = {}

        print "transaction listing"

        ndto = datetime.datetime.strptime(dto, "%Y-%m-%d")
        todate = datetime.date(int(ndto.year), int(ndto.month), 10)
        toyear = todate.year
        tomonth = todate.month
        nfrom = datetime.datetime.strptime(dfrom, "%Y-%m-%d")
        fromdate = datetime.date(int(nfrom.year), int(nfrom.month), 10)
        fromyear = fromdate.year
        frommonth = fromdate.month

        budget = Departmentbudget.objects.filter(isdeleted=0, year=toyear)

        if chart != '':
            budget = budget.filter(chartofaccount_id=chart)
            chartofaccount = Chartofaccount.objects.filter(isdeleted=0, id__exact=chart).first()
            print 'budget'

        if department != '':
            budget = budget.filter(department_id=department)
            dept = Department.objects.filter(isdeleted=0, id__exact=department).first()
            print 'department'

        # budget = getBudget(tomonth, item)
        totalbudget = 0
        for item in budget:
            budgetdata = getBudget(tomonth, item)
            totalbudget += budgetdata[1]

        title = "Department Budget"

        result = query_transaction(dto, dfrom, chart, department)


        # Create an in-memory output file for the new workbook.
        output = io.BytesIO()

        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()

        # variables
        bold = workbook.add_format({'bold': 1})
        formatdate = workbook.add_format({'num_format': 'yyyy/mm/dd'})
        centertext = workbook.add_format({'bold': 1, 'align': 'center'})

        # title
        worksheet.write('A1', 'DEPARTMENT BUDGET', bold)
        worksheet.write('A2', 'AS OF '+str(dfrom)+' to '+str(dto), bold)

        if chartofaccount and  dept:
            worksheet.write('A3', 'Chart of Account', bold)
            worksheet.write('B3', chartofaccount.accountcode, bold)
            worksheet.write('C3', chartofaccount.description, bold)
            worksheet.write('A4', 'Department', bold)
            worksheet.write('B4', dept.code, bold)
            worksheet.write('C4', dept.departmentname, bold)
        elif chartofaccount:
            worksheet.write('A3', 'Chart of Account', bold)
            worksheet.write('B3', chartofaccount.accountcode, bold)
            worksheet.write('C3', chartofaccount.description, bold)
        elif dept:
            worksheet.write('A3', 'Department', bold)
            worksheet.write('B3', dept.code, bold)
            worksheet.write('C3', dept.departmentname, bold)
        else:
            worksheet.write('A3', 'ALL Transaction', bold)

        # header
        worksheet.write('A6', 'Date', bold)
        worksheet.write('B6', 'Type', bold)
        worksheet.write('C6', 'Number', bold)
        worksheet.write('D6', 'Dept', bold)
        worksheet.write('E6', 'Particulars', bold)
        worksheet.write('F6', 'Debit Amount', bold)
        worksheet.write('G6', 'Credit Amount', bold)

        row = 7
        col = 0
        debit = 0
        credit = 0
        total = 0
        totalvariance = 0

        for data in result:
            debit += data.debitamount
            credit += data.creditamount
            worksheet.write(row, col, data.tdate, formatdate)
            worksheet.write(row, col + 1, data.tran)
            worksheet.write(row, col + 2, data.tnum)
            worksheet.write(row, col + 3, data.deptcode)
            worksheet.write(row, col + 4, data.particulars)
            worksheet.write(row, col + 5, float(format(data.debitamount, '.2f')))
            worksheet.write(row, col + 6, float(format(data.creditamount, '.2f')))

            row += 1

        total = float(debit) - float(credit)
        totalvariance = float(totalbudget) - float(total)

        worksheet.write(row, col + 4, 'Total', bold)
        worksheet.write(row, col + 5, float(format(debit, '.2f')), bold)
        worksheet.write(row, col + 6, float(format(credit, '.2f')), bold)

        worksheet.write(row+1, col + 4, 'NET Amount', bold)
        if debit > credit:
            worksheet.write(row+1, col + 5, float(format(total, '.2f')), bold)
            worksheet.write(row+1, col + 6, float(format(0, '.2f')), bold)
        else:
            worksheet.write(row + 1, col + 5, float(format(0, '.2f')), bold)
            worksheet.write(row + 1, col + 6, float(format(total, '.2f')), bold)

        worksheet.write('A5', 'Budget Amount', bold)
        worksheet.write('B5', float(format(totalbudget, '.2f')), bold)
        worksheet.write('C5', 'Actual Amount', bold)
        worksheet.write('D5', float(format(total, '.2f')), bold)
        worksheet.write('E5', 'Variance Amount', bold)
        worksheet.write('F5', float(format(totalvariance, '.2f')), bold)

        workbook.close()

        # Rewind the buffer.
        output.seek(0)

        # Set up the Http response.
        filename = "departmentbudget.xlsx"
        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename=%s' % filename

        return response

class GeneratePDF(View):
    def get(self, request):
        company = Companyparameter.objects.all().first()
        list = []
        total = []
        report = request.GET['report']
        filter = request.GET['filter']
        type = request.GET['type']
        dfrom = request.GET['from']
        dto = request.GET['to']
        department = request.GET['department']
        product = request.GET['product']
        title = "Budget Monitoring Report"
        subtitle = ""
        typetext = "Summary"
        filtertext = "Cost of Sales"

        ndto = datetime.datetime.strptime(dto, "%Y-%m-%d")
        todate = datetime.date(int(ndto.year), int(ndto.month), 10)
        toyear = todate.year
        tomonth = todate.month
        nfrom = datetime.datetime.strptime(dfrom, "%Y-%m-%d")
        fromdate = datetime.date(int(nfrom.year), int(nfrom.month), 10)
        fromyear = fromdate.year
        frommonth = fromdate.month

        prevdate = datetime.date(int(fromyear), int(frommonth), 10) - timedelta(days=15)
        prevyear = prevdate.year
        prevmonth = prevdate.month

        if filter == '527':
            filtertext = "Cost of Sales"
        elif filter == '643':
            filtertext = "General & Administrative"
        elif filter == '745':
            filtertext = "Selling Expense"
        else:
            filtertext = "All"

        if type == '1':
            typetext = "Summary"
        else:
            typetext = "Detailed"

        list = []

        if report == '1':
            if type == '2':
                list = query_department_budget(type, filter, department, product, fromyear, frommonth, toyear, tomonth, prevyear, prevmonth)
                title = "Budget Report - Department/Section - Detailed"
            else:
                list = query_department_budget(type, filter, department, product, fromyear, frommonth, toyear, tomonth, prevyear, prevmonth)
                title = "Budget Report - Department/Section - Summary"

        context = {
            "title": title,
            "subtitle": subtitle,
            "asof": ndto,
            "prevdate": prevdate,
            "curdate": nfrom,
            "company": company,
            "list": list,
            "filtertext": filtertext,
            "typetext": typetext,
            "total": total,
            "username": request.user,
        }

        if report == '1':
            if type == '2':
                print 'detail'
                return Render.render('budgetreport/department_detail.html', context)
            else:
                print 'summary'
                return Render.render('budgetreport/department_summary.html', context)
        else:
            return Render.render('budgetreport/department_summary.html', context)

@method_decorator(login_required, name='dispatch')
class GenerateExcel(View):
    def get(self, request):
        company = Companyparameter.objects.all().first()
        list = []
        total = []
        report = request.GET['report']
        filter = request.GET['filter']
        type = request.GET['type']
        dfrom = request.GET['from']
        dto = request.GET['to']
        department = request.GET['department']
        product = request.GET['product']
        title = "Budget Monitoring Report"
        subtitle = ""
        typetext = "Summary"
        filtertext = "Cost of Sales"

        ndto = datetime.datetime.strptime(dto, "%Y-%m-%d")

        todate = datetime.date(int(ndto.year), int(ndto.month), 10)
        toyear = todate.year
        tomonth = todate.month
        nfrom = datetime.datetime.strptime(dfrom, "%Y-%m-%d")
        nfrommonth = nfrom.strftime("%B")
        fromdate = datetime.date(int(nfrom.year), int(nfrom.month), 10)
        fromyear = fromdate.year
        frommonth = fromdate.month

        prevdate = datetime.date(int(fromyear), int(frommonth), 10) - timedelta(days=15)
        prevyear = prevdate.year
        prevmonth = prevdate.month
        prevmon = prevdate.strftime("%B")

        if filter == '527':
            filtertext = "Cost of Sales"
        elif filter == '643':
            filtertext = "General & Administrative"
        elif filter == '745':
            filtertext = "Selling Expense"
        else:
            filtertext = "All"

        if type == '1':
            typetext = "Summary"
        else:
            typetext = "Detailed"

        list = []
        filename = "departmentbudget.xlsx"

        if report == '1':
            if type == '2':
                list = query_department_budget(type, filter, department, product, fromyear, frommonth, toyear, tomonth,
                                               prevyear, prevmonth)
                title = "Budget Report - Department/Section - Detailed"
                filename = "budgetreport-department-detailed.xlsx"
            else:
                list = query_department_budget(type, filter, department, product, fromyear, frommonth, toyear, tomonth,
                                               prevyear, prevmonth)
                title = "Budget Report - Department/Section - Summary"
                filename = "budgetreport-department-summary.xlsx"

        # Create an in-memory output file for the new workbook.
        output = io.BytesIO()

        workbook = xlsxwriter.Workbook(output)
        worksheet = workbook.add_worksheet()

        # variables
        bold = workbook.add_format({'bold': 1})
        formatdate = workbook.add_format({'num_format': 'yyyy/mm/dd'})
        centertext = workbook.add_format({'bold': 1, 'align': 'center'})

        # title
        worksheet.write('A1', str(title), bold)
        worksheet.write('A2', str(filtertext), bold)
        worksheet.write('A3', 'for the period ending '+str(dto), bold)

        df = pd.DataFrame(list)
        if report == '1':

            # header
            worksheet.merge_range   ('B5:D5', '-------------------- actual --------------------', bold)

            worksheet.write('A6', '', bold)
            worksheet.write('B6', str(nfrommonth), bold)
            worksheet.write('C6', str(prevmon), bold)
            worksheet.write('D6', 'year-to-date', bold)
            worksheet.write('E6', 'variance over/(under)', bold)
            worksheet.write('F6', 'variance(%) over/(under)', bold)

            row = 6
            col = 0
            if type == '1':
                print 'summary'
                for dept, department in df.fillna('NaN').groupby(['dcode', 'departmentname']):

                    worksheet.write(row, col, str(dept[0])+'-'+str(dept[1]), bold)
                    row += 1

                    totalcuramount = 0
                    totalprevamount = 0
                    totalytdamount = 0
                    totalvaramount = 0
                    totalvarpercent = 0
                    for group, subgroup in department.fillna('NaN').groupby(['cgrouptitle', 'cgroupdescription']):
                        worksheet.write(row, col, str(group[0]), bold)
                        row += 1
                        for head, headgroup in subgroup.fillna('NaN').groupby(['csubgrouptitle', 'csubgroupdescription']):
                            worksheet.write(row, col, '  '+str(head[0]), bold)
                            row += 1
                            subcuramount = 0
                            subprevamount = 0
                            subytdamount = 0
                            subvaramount = 0
                            subvarpercent = 0
                            for data, item in headgroup.iterrows():
                                worksheet.write(row, col, '   '+str(item.csubheaddescription))
                                worksheet.write(row, col + 1, float(format(item.curamount, '.2f')))
                                worksheet.write(row, col + 2, float(format(item.prevamount, '.2f')))
                                worksheet.write(row, col + 3, float(format(item.ytdamount, '.2f')))
                                worksheet.write(row, col + 4, float(format(item.varamount, '.2f')))
                                worksheet.write(row, col + 5, float(format(item.varpercent, '.2f')))
                                subcuramount += item.curamount
                                subprevamount += item.prevamount
                                subytdamount += item.ytdamount
                                subvaramount += item.varamount
                                row += 1

                            if subprevamount > 0:
                                subvarpercent = (subvaramount/subprevamount) * 100

                            worksheet.write(row, col, '  Subtotal - '+str(head[0]))
                            worksheet.write(row, col + 1, float(format(subcuramount, '.2f')))
                            worksheet.write(row, col + 2, float(format(subprevamount, '.2f')))
                            worksheet.write(row, col + 3, float(format(subytdamount, '.2f')))
                            worksheet.write(row, col + 4, float(format(subvaramount, '.2f')))
                            worksheet.write(row, col + 5, float(format(subvarpercent, '.2f')))

                            totalcuramount += subcuramount
                            totalprevamount += subprevamount
                            totalytdamount += subytdamount
                            totalvaramount += subvaramount

                            row += 1

                    if totalprevamount > 0:
                        totalvarpercent = (totalvaramount / totalprevamount) * 100

                    worksheet.write(row, col, 'Total - ' + str(group[0]))
                    worksheet.write(row, col + 1, float(format(totalcuramount, '.2f')))
                    worksheet.write(row, col + 2, float(format(totalprevamount, '.2f')))
                    worksheet.write(row, col + 3, float(format(totalytdamount, '.2f')))
                    worksheet.write(row, col + 4, float(format(totalvaramount, '.2f')))
                    worksheet.write(row, col + 5, float(format(totalvarpercent, '.2f')))
                    row += 1
                    row += 1
            else:
                print 'detailed'
                for dept, department in df.fillna('NaN').groupby(['dcode', 'departmentname']):

                    worksheet.write(row, col, str(dept[0])+'-'+str(dept[1]), bold)
                    row += 1

                    totalcuramount = 0
                    totalprevamount = 0
                    totalytdamount = 0
                    totalvaramount = 0
                    totalvarpercent = 0
                    for group, subgroup in department.fillna('NaN').groupby(['cgrouptitle', 'cgroupdescription']):
                        worksheet.write(row, col, str(group[0]), bold)
                        row += 1

                        for head, headgroup in subgroup.fillna('NaN').groupby(['csubgrouptitle', 'csubgroupdescription']):
                            worksheet.write(row, col, '  '+str(head[0]), bold)
                            row += 1

                            subcuramount = 0
                            subprevamount = 0
                            subytdamount = 0
                            subvaramount = 0
                            subvarpercent = 0
                            for subhead, subheadgroup in headgroup.fillna('NaN').groupby(['csubheadtitle', 'csubheaddescription']):
                                worksheet.write(row, col, '    ' + str(subhead[0]), bold)
                                row += 1
                                for data, item in subheadgroup.iterrows():
                                    worksheet.write(row, col, '    ' + str(item.description))
                                    worksheet.write(row, col + 1, float(format(item.curamount, '.2f')))
                                    worksheet.write(row, col + 2, float(format(item.prevamount, '.2f')))
                                    worksheet.write(row, col + 3, float(format(item.ytdamount, '.2f')))
                                    worksheet.write(row, col + 4, float(format(item.varamount, '.2f')))
                                    worksheet.write(row, col + 5, float(format(item.varpercent, '.2f')))
                                    subcuramount += item.curamount
                                    subprevamount += item.prevamount
                                    subytdamount += item.ytdamount
                                    subvaramount += item.varamount
                                    row += 1

                            if subprevamount > 0:
                                subvarpercent = (subvaramount/subprevamount) * 100

                            worksheet.write(row, col, '  Subtotal - ' + str(head[0]))
                            worksheet.write(row, col + 1, float(format(subcuramount, '.2f')))
                            worksheet.write(row, col + 2, float(format(subprevamount, '.2f')))
                            worksheet.write(row, col + 3, float(format(subytdamount, '.2f')))
                            worksheet.write(row, col + 4, float(format(subvaramount, '.2f')))
                            worksheet.write(row, col + 5, float(format(subvarpercent, '.2f')))

                            totalcuramount += subcuramount
                            totalprevamount += subprevamount
                            totalytdamount += subytdamount
                            totalvaramount += subvaramount
                            row += 1

                        if totalprevamount > 0:
                            totalvarpercent = (totalvaramount / totalprevamount) * 100

                        worksheet.write(row, col, '  Total - ' + str(group[0]))
                        worksheet.write(row, col + 1, float(format(totalcuramount, '.2f')))
                        worksheet.write(row, col + 2, float(format(totalprevamount, '.2f')))
                        worksheet.write(row, col + 3, float(format(totalytdamount, '.2f')))
                        worksheet.write(row, col + 4, float(format(totalvaramount, '.2f')))
                        worksheet.write(row, col + 5, float(format(totalvarpercent, '.2f')))
                        row += 1
                        row += 1

        workbook.close()

        # Rewind the buffer.
        output.seek(0)

        # Set up the Http response.
        #filename = "departmentbudget.xlsx"
        response = HttpResponse(
            output,
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        response['Content-Disposition'] = 'attachment; filename=%s' % filename

        return response

def query_department_budget(type, expense, department, product, fromyear, frommonth, toyear, tomonth, prevyear, prevmonth):
    print "Transaction Query Department Budget"
    ''' Create query '''
    cursor = connection.cursor()

    type_condition = ''
    department_condition = ''
    expense_condition = ''
    product_condition = ''

    if type == '2':
        type_condition = "GROUP BY d.code, z.description"
    else:
        type_condition = "GROUP BY d.code, z.csubheadtitle"

    if expense != '':
        expense_condition = "AND d.expchartofaccount_id = '" + str(expense) + "'"

    if product != '':
        product_condition = "AND d.product_id = '" + str(product) + "'"

    if department != '':
        department_condition = "AND a.department_id = '" + str(department) + "'"
        #department_condition = "AND a.department_id IN (96, 3)"

    print 'start'
    query = "SELECT d.code AS dcode, d.departmentname, d.product_id, d.expchartofaccount_id, SUM(IF(z.code = 'C', z.amount * -1, z.amount)) AS curamount, SUM(IF(z.code = 'C', z.prevamount * -1, z.prevamount)) AS prevamount, SUM(z.ytdamount) AS ytdamount, " \
            "(SUM(IF(z.code = 'C', z.amount * -1, z.amount)) - SUM(IF(z.code = 'C', z.prevamount * -1, z.prevamount))) AS varamount, " \
            "IFNULL(ROUND(((SUM(IF(z.code = 'C', z.amount * -1, z.amount)) - SUM(IF(z.code = 'C', z.prevamount * -1, z.prevamount))) / SUM(IF(z.code = 'C', z.prevamount * -1, z.prevamount))) * 100, 2), 0) AS varpercent, " \
            "z.accountcode, z.title, z.description, " \
            "z.cgrouptitle, z.cgroupdescription, z.csubgrouptitle, z.csubgroupdescription, z.csubheadtitle, z.csubheaddescription " \
            "FROM ( " \
            "SELECT a.year, a.month, a.amount, 0 AS prevamount, 0 AS ytdamount, a.code, a.chartofaccount_id, a.department_id, " \
            "c.main, c.clas, c.item, c.cont, c.sub, c.accountcode, c.title, c.description, c.accounttype, " \
            "cgroup.title AS cgrouptitle, cgroup.description AS cgroupdescription, " \
            "csubgroup.title AS csubgrouptitle, csubgroup.description AS csubgroupdescription, " \
            "csubhead.title AS csubheadtitle, csubhead.description AS csubheaddescription " \
            "FROM accountexpensebalance AS a " \
            "LEFT OUTER JOIN chartofaccount AS c ON c.id = a.chartofaccount_id " \
            "LEFT OUTER JOIN chartofaccount AS cgroup ON (cgroup.main = c.main AND cgroup.clas = c.clas AND cgroup.item = 0 AND cgroup.cont = 0 AND cgroup.sub = 000000 AND cgroup.accounttype = 'T') " \
            "LEFT OUTER JOIN chartofaccount AS csubgroup ON (csubgroup.main = c.main AND csubgroup.clas = c.clas AND csubgroup.item = c.item AND csubgroup.cont = 0 AND csubgroup.sub = 000000 AND csubgroup.accounttype = 'T') " \
            "LEFT OUTER JOIN chartofaccount AS csubhead ON (csubhead.main = c.main AND csubhead.clas = c.clas AND csubhead.item = c.item AND csubhead.cont = c.cont AND csubhead.sub = CONCAT(SUBSTR(c.sub, 1, 1),'','00000')) " \
            "WHERE a.year = "+str(fromyear)+" AND a.month = "+str(frommonth)+" "+str(department_condition)+" " \
            "UNION " \
            "SELECT a.year, a.month, 0 AS curamount, a.amount, 0 AS ytdamount, a.code, a.chartofaccount_id, a.department_id, " \
            "c.main, c.clas, c.item, c.cont, c.sub, " \
            "c.accountcode, c.title, c.description, c.accounttype, " \
            "cgroup.title AS cgrouptitle, cgroup.description AS cgroupdescription, " \
            "csubgroup.title AS csubgrouptitle, csubgroup.description AS csubgroupdescription, " \
            "csubhead.title AS csubheadtitle, csubhead.description AS csubheaddescription " \
            "FROM accountexpensebalance AS a " \
            "LEFT OUTER JOIN chartofaccount AS c ON c.id = a.chartofaccount_id " \
            "LEFT OUTER JOIN chartofaccount AS cgroup ON (cgroup.main = c.main AND cgroup.clas = c.clas AND cgroup.item = 0 AND cgroup.cont = 0 AND cgroup.sub = 000000 AND cgroup.accounttype = 'T') " \
            "LEFT OUTER JOIN chartofaccount AS csubgroup ON (csubgroup.main = c.main AND csubgroup.clas = c.clas AND csubgroup.item = c.item AND csubgroup.cont = 0 AND csubgroup.sub = 000000 AND csubgroup.accounttype = 'T') " \
            "LEFT OUTER JOIN chartofaccount AS csubhead ON (csubhead.main = c.main AND csubhead.clas = c.clas AND csubhead.item = c.item AND csubhead.cont = c.cont AND csubhead.sub = CONCAT(SUBSTR(c.sub, 1, 1),'','00000')) " \
            "WHERE a.year = "+str(prevyear)+" AND a.month = "+str(prevmonth)+" "+str(department_condition)+" " \
            "UNION " \
            "SELECT a.year, a.month, 0 AS curamount, 0 AS amount, SUM(IF(a.code = 'C', a.amount * -1, a.amount)) AS ytdamount,a.code, a.chartofaccount_id, a.department_id, " \
            "c.main, c.clas, c.item, c.cont, c.sub, " \
            "c.accountcode, c.title, c.description, c.accounttype, " \
            "cgroup.title AS cgrouptitle, cgroup.description AS cgroupdescription, " \
            "csubgroup.title AS csubgrouptitle, csubgroup.description AS csubgroupdescription, " \
            "csubhead.title AS csubheadtitle, csubhead.description AS csubheaddescription " \
            "FROM accountexpensebalance AS a " \
            "LEFT OUTER JOIN chartofaccount AS c ON c.id = a.chartofaccount_id " \
            "LEFT OUTER JOIN chartofaccount AS cgroup ON (cgroup.main = c.main AND cgroup.clas = c.clas AND cgroup.item = 0 AND cgroup.cont = 0 AND cgroup.sub = 000000 AND cgroup.accounttype = 'T') " \
            "LEFT OUTER JOIN chartofaccount AS csubgroup ON (csubgroup.main = c.main AND csubgroup.clas = c.clas AND csubgroup.item = c.item AND csubgroup.cont = 0 AND csubgroup.sub = 000000 AND csubgroup.accounttype = 'T') " \
            "LEFT OUTER JOIN chartofaccount AS csubhead ON (csubhead.main = c.main AND csubhead.clas = c.clas AND csubhead.item = c.item AND csubhead.cont = c.cont AND csubhead.sub = CONCAT(SUBSTR(c.sub, 1, 1),'','00000')) " \
            "WHERE a.year = "+str(fromyear)+" AND a.month >= 1 AND a.month <= "+str(tomonth)+" "+str(department_condition)+" " \
            "GROUP BY c.accountcode " \
            ") AS z " \
            "LEFT OUTER JOIN department AS d ON d.id = z.department_id " \
            "WHERE d.isdeleted = 0 "+str(expense_condition)+" "+str(product_condition)+" "+" "+str(type_condition)+" " \
            "ORDER BY d.code, z.accountcode, z.year ASC, z.month DESC"

    cursor.execute(query)
    result = namedtuplefetchall(cursor)
    print 'end'
    return result