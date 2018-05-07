import datetime
from django.views.generic import ListView, DetailView, CreateView, UpdateView, DeleteView
from django.contrib.auth.decorators import login_required
from django.utils.decorators import method_decorator
from django.http import HttpResponseRedirect, Http404
from departmentbudget.models import Departmentbudget
from department.models import Department
from unit.models import Unit
from chartofaccount.models import Chartofaccount
from django.views.decorators.csrf import csrf_exempt
from django.db.models import Q, Sum
from annoying.functions import get_object_or_None


@method_decorator(login_required, name='dispatch')
class IndexView(ListView):
    model = Departmentbudget
    template_name = 'departmentbudget/index.html'
    context_object_name = 'data_list'

    def get_queryset(self):
        return Departmentbudget.objects.all().filter(isdeleted=0).order_by('-pk')


@method_decorator(login_required, name='dispatch')
class DetailView(DetailView):
    model = Departmentbudget
    template_name = 'departmentbudget/detail.html'


@method_decorator(login_required, name='dispatch')
class CreateView(CreateView):
    model = Departmentbudget
    template_name = 'departmentbudget/create.html'
    fields = ['year', 'department', 'unit', 'chartofaccount',
              'remarks', 'formula', 'method',
              'mjan', 'mfeb', 'mmar',
              'mapr', 'mmay', 'mjun',
              'mjul', 'maug', 'msep',
              'moct', 'mnov', 'mdec']

    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('departmentbudget.add_departmentbudget'):
            raise Http404
        return super(CreateView, self).dispatch(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        context = super(CreateView, self).get_context_data(**kwargs)
        context['unit'] = Unit.objects.filter(isdeleted=0).order_by('description')
        if self.request.POST.get('department', False):
            context['department'] = Department.objects.get(pk=self.request.POST['department'], isdeleted=0)
        if self.request.POST.get('chartofaccount', False):
            context['chartofaccount'] = Chartofaccount.objects.get(pk=self.request.POST['chartofaccount'], isdeleted=0)
        return context

    def form_valid(self, form):
        self.object = form.save(commit=False)
        self.object.enterby = self.request.user
        self.object.modifyby = self.request.user
        self.object.save()
        return HttpResponseRedirect('/departmentbudget')


@method_decorator(login_required, name='dispatch')
class UpdateView(UpdateView):
    model = Departmentbudget
    template_name = 'departmentbudget/edit.html'
    fields = ['year', 'department', 'unit', 'chartofaccount',
              'remarks', 'formula', 'method',
              'mjan', 'mfeb', 'mmar',
              'mapr', 'mmay', 'mjun',
              'mjul', 'maug', 'msep',
              'moct', 'mnov', 'mdec']

    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('departmentbudget.change_departmentbudget'):
            raise Http404
        return super(UpdateView, self).dispatch(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        context = super(UpdateView, self).get_context_data(**kwargs)
        context['unit'] = Unit.objects.filter(isdeleted=0).order_by('description')
        if self.request.POST.get('chartofaccount', False):
            context['chartofaccount'] = Chartofaccount.objects.get(pk=self.request.POST['chartofaccount'], isdeleted=0)
        elif self.object.chartofaccount:
            context['chartofaccount'] = Chartofaccount.objects.get(pk=self.object.chartofaccount.id, isdeleted=0)
        if self.request.POST.get('department', False):
            context['department'] = Department.objects.get(pk=self.request.POST['department'], isdeleted=0)
        elif self.object.department:
            context['department'] = Department.objects.get(pk=self.object.department.id, isdeleted=0)
        return context

    def form_valid(self, form):
        self.object = form.save(commit=False)
        self.object.modifyby = self.request.user
        self.object.modifydate = datetime.datetime.now()
        self.object.save()
        return HttpResponseRedirect('/departmentbudget')


@method_decorator(login_required, name='dispatch')
class DeleteView(DeleteView):
    model = Departmentbudget
    template_name = 'departmentbudget/delete.html'

    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('departmentbudget.delete_departmentbudget'):
            raise Http404
        return super(DeleteView, self).dispatch(request, *args, **kwargs)

    def delete(self, request, *args, **kwargs):
        self.object = self.get_object()
        self.object.modifyby = self.request.user
        self.object.modifydate = datetime.datetime.now()
        self.object.isdeleted = 1
        self.object.status = 'I'
        self.object.save()
        return HttpResponseRedirect('/departmentbudget')


@method_decorator(login_required, name='dispatch')
class ReportView(ListView):
    model = Departmentbudget
    template_name = 'departmentbudget/report.html'

    def get_context_data(self, **kwargs):
        context = super(ListView, self).get_context_data(**kwargs)

        context['department'] = Department.objects.filter(isdeleted=0).order_by('code')

        return context


@method_decorator(login_required, name='dispatch')
class ReportResultHtmlView(ListView):
    model = Departmentbudget
    template_name = 'departmentbudget/reportresulthtml.html'

    def get_context_data(self, **kwargs):
        context = super(ListView, self).get_context_data(**kwargs)
        context['report_type'] = ''
        context['report_total'] = 0

        query, context['report_type'], context['report_total'], context['report_xls'] = reportresultquery(self.request)

        context['report'] = self.request.COOKIES.get('rep_f_group_' + self.request.resolver_match.app_name)
        context['data_list'] = query

        # pdf config
        context['rc_orientation'] = ('portrait', 'landscape')[self.request.COOKIES.get('rep_f_orientation_' + self.request.resolver_match.app_name) == 'l']
        context['rc_headtitle'] = "DEPARTMENT BUDGET"
        context['rc_title'] = "DEPARTMENT BUDGET"

        return context


@csrf_exempt
def reportresultquery(request):
    query = ''
    report_type = ''
    report_xls = ''
    report_total = ''

    query = Departmentbudget.objects.all().filter(isdeleted=0)

    if request.COOKIES.get('rep_f_year_' + request.resolver_match.app_name):
        key_data = str(request.COOKIES.get('rep_f_year_' + request.resolver_match.app_name))
        query = query.filter(year=key_data)
    if request.COOKIES.get('rep_f_department_' + request.resolver_match.app_name):
        key_data = str(request.COOKIES.get('rep_f_department_' + request.resolver_match.app_name))
        query = query.filter(department=int(key_data))
    if request.COOKIES.get('rep_f_gl_' + request.resolver_match.app_name) is not None \
            and request.COOKIES.get('rep_f_gl_' + request.resolver_match.app_name) != 'null':
        key_data = str(request.COOKIES.get('rep_f_gl_' + request.resolver_match.app_name))
        query = query.filter(chartofaccount=get_object_or_None(Chartofaccount, pk=int(key_data)))

    if request.COOKIES.get('rep_f_group_' + request.resolver_match.app_name) == 'ds':

        report_type = "Department Budget (Department Summary)"
        report_xls = "Dept. Budget (Dept. Summary)"

        query = query.values('department', 'department__departmentname', 'department__code')\
                     .annotate(
                               Sum('mjan'),
                               Sum('mfeb'),
                               Sum('mmar'),
                               Sum('mapr'),
                               Sum('mmay'),
                               Sum('mjun'),
                               Sum('mjul'),
                               Sum('maug'),
                               Sum('msep'),
                               Sum('moct'),
                               Sum('mnov'),
                               Sum('mdec'),
                               total=Sum('mjan')+Sum('mfeb')+Sum('mmar')+Sum('mapr')+Sum('mmay')+Sum('mjun')+Sum('mjul')+Sum('maug')+Sum('msep')+Sum('moct')+Sum('mnov')+Sum('mdec'),
                               )\
                     .order_by('department__code')\

        report_total = query.aggregate(
                                        Sum('mjan__sum'),
                                        Sum('mfeb__sum'),
                                        Sum('mmar__sum'),
                                        Sum('mapr__sum'),
                                        Sum('mmay__sum'),
                                        Sum('mjun__sum'),
                                        Sum('mjul__sum'),
                                        Sum('maug__sum'),
                                        Sum('msep__sum'),
                                        Sum('moct__sum'),
                                        Sum('mnov__sum'),
                                        Sum('mdec__sum'),
                                        Sum('total'),
                                      )
    elif request.COOKIES.get('rep_f_group_' + request.resolver_match.app_name) == 'as':

        report_type = "Department Budget (Account Summary)"
        report_xls = "Dept. Budget (Acct. Summary)"

        query = query.values('chartofaccount', 'chartofaccount__accountcode', 'chartofaccount__description')\
                     .annotate(
                               Sum('mjan'),
                               Sum('mfeb'),
                               Sum('mmar'),
                               Sum('mapr'),
                               Sum('mmay'),
                               Sum('mjun'),
                               Sum('mjul'),
                               Sum('maug'),
                               Sum('msep'),
                               Sum('moct'),
                               Sum('mnov'),
                               Sum('mdec'),
                               total=Sum('mjan')+Sum('mfeb')+Sum('mmar')+Sum('mapr')+Sum('mmay')+Sum('mjun')+Sum('mjul')+Sum('maug')+Sum('msep')+Sum('moct')+Sum('mnov')+Sum('mdec'),
                               )\
                     .order_by('chartofaccount__accountcode')\

        report_total = query.aggregate(
                                        Sum('mjan__sum'),
                                        Sum('mfeb__sum'),
                                        Sum('mmar__sum'),
                                        Sum('mapr__sum'),
                                        Sum('mmay__sum'),
                                        Sum('mjun__sum'),
                                        Sum('mjul__sum'),
                                        Sum('maug__sum'),
                                        Sum('msep__sum'),
                                        Sum('moct__sum'),
                                        Sum('mnov__sum'),
                                        Sum('mdec__sum'),
                                        Sum('total'),
                                      )
    elif request.COOKIES.get('rep_f_group_' + request.resolver_match.app_name) == 'dd':

        report_type = "Department Budget (Department Detailed)"
        report_xls = "Dept. Budget (Dept. Detailed)"

        query = query.values('department', 'department__departmentname', 'department__code', 'chartofaccount', 'chartofaccount__accountcode', 'chartofaccount__description')\
                     .annotate(
                               Sum('mjan'),
                               Sum('mfeb'),
                               Sum('mmar'),
                               Sum('mapr'),
                               Sum('mmay'),
                               Sum('mjun'),
                               Sum('mjul'),
                               Sum('maug'),
                               Sum('msep'),
                               Sum('moct'),
                               Sum('mnov'),
                               Sum('mdec'),
                               total=Sum('mjan')+Sum('mfeb')+Sum('mmar')+Sum('mapr')+Sum('mmay')+Sum('mjun')+Sum('mjul')+Sum('maug')+Sum('msep')+Sum('moct')+Sum('mnov')+Sum('mdec'),
                               )\
                     .order_by('department__code', 'chartofaccount__accountcode')\

        report_total = query.aggregate(
                                        Sum('mjan__sum'),
                                        Sum('mfeb__sum'),
                                        Sum('mmar__sum'),
                                        Sum('mapr__sum'),
                                        Sum('mmay__sum'),
                                        Sum('mjun__sum'),
                                        Sum('mjul__sum'),
                                        Sum('maug__sum'),
                                        Sum('msep__sum'),
                                        Sum('moct__sum'),
                                        Sum('mnov__sum'),
                                        Sum('mdec__sum'),
                                        Sum('total'),
                                      )
    elif request.COOKIES.get('rep_f_group_' + request.resolver_match.app_name) == 'ad':

        report_type = "Department Budget (Account Detailed)"
        report_xls = "Dept. Budget (Acct Detailed)"

        query = query.values('chartofaccount', 'chartofaccount__accountcode', 'chartofaccount__description', 'department', 'department__departmentname', 'department__code')\
                     .annotate(
                               Sum('mjan'),
                               Sum('mfeb'),
                               Sum('mmar'),
                               Sum('mapr'),
                               Sum('mmay'),
                               Sum('mjun'),
                               Sum('mjul'),
                               Sum('maug'),
                               Sum('msep'),
                               Sum('moct'),
                               Sum('mnov'),
                               Sum('mdec'),
                               total=Sum('mjan')+Sum('mfeb')+Sum('mmar')+Sum('mapr')+Sum('mmay')+Sum('mjun')+Sum('mjul')+Sum('maug')+Sum('msep')+Sum('moct')+Sum('mnov')+Sum('mdec'),
                               )\
                     .order_by('chartofaccount__accountcode', 'department__code')\

        report_total = query.aggregate(
                                        Sum('mjan__sum'),
                                        Sum('mfeb__sum'),
                                        Sum('mmar__sum'),
                                        Sum('mapr__sum'),
                                        Sum('mmay__sum'),
                                        Sum('mjun__sum'),
                                        Sum('mjul__sum'),
                                        Sum('maug__sum'),
                                        Sum('msep__sum'),
                                        Sum('moct__sum'),
                                        Sum('mnov__sum'),
                                        Sum('mdec__sum'),
                                        Sum('total'),
                                      )

    return query, report_type, report_total, report_xls
