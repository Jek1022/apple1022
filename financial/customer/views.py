import datetime
from django.views.generic import View, ListView, DetailView, CreateView, UpdateView, DeleteView
from django.contrib.auth.decorators import login_required
from django.utils.decorators import method_decorator
from django.http import HttpResponseRedirect, Http404
from customertype.models import Customertype
from creditterm.models import Creditterm
from currency.models import Currency
from bankaccount.models import Bankaccount
from industry.models import Industry
from vat.models import Vat
from . models import Customer
from financial.utils import Render
from django.utils import timezone
from companyparameter.models import Companyparameter

# pagination and search
from endless_pagination.views import AjaxListView
from django.db.models import Q

# Create your views here.
@method_decorator(login_required, name='dispatch')
class IndexView(AjaxListView):
    model = Customer
    template_name = 'customer/index.html'
    context_object_name = 'data_list'

    # pagination and search
    page_template = 'customer/index_list.html'
    def get_queryset(self):
        query = Customer.objects.all().filter(isdeleted=0)
        if self.request.COOKIES.get('keysearch_' + self.request.resolver_match.app_name):
            keysearch = str(self.request.COOKIES.get('keysearch_' + self.request.resolver_match.app_name))
            query = query.filter(Q(code__icontains=keysearch) |
                                 Q(name__icontains=keysearch))
        return query


@method_decorator(login_required, name='dispatch')
class DetailView(DetailView):
    model = Customer
    template_name = 'customer/detail.html'


@method_decorator(login_required, name='dispatch')
class CreateView(CreateView):
    model = Customer
    template_name = 'customer/create.html'
    fields = ['code', 'name', 'address1', 'address2', 'address3', 'telno1',
              'telno2', 'telno3', 'faxno1', 'faxno2',
              'tin', 'pagerno', 'payterms', 'creditlimit',
              'creditstatus', 'creditrating', 'contactperson',
              'contactposition', 'contactemail', 'remarks',
              'beg_amount', 'beg_code', 'beg_date',
              'end_amount', 'end_code', 'end_date', 'bankaccount',
              'creditterm', 'currency', 'customertype', 'industry', 'vat']

    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('customer.add_customer'):
            raise Http404
        return super(CreateView, self).dispatch(request, *args, **kwargs)

    def form_valid(self, form):
        self.object = form.save(commit=False)
        self.object.multiplestatus = 'N'
        self.object.enterby = self.request.user
        self.object.modifyby = self.request.user

        if self.object.vat:
            self.object.vatrate = self.object.vat.rate

        self.object.save()
        return HttpResponseRedirect('/customer')

    def get_context_data(self, **kwargs):
        context = super(CreateView, self).get_context_data(**kwargs)
        context['customertype'] = Customertype.objects.filter(isdeleted=0).order_by('description')
        context['creditterm'] = Creditterm.objects.filter(isdeleted=0).order_by('description')
        context['currency'] = Currency.objects.filter(isdeleted=0).order_by('description')
        context['bankaccount'] = Bankaccount.objects.filter(isdeleted=0).order_by('bank')
        context['industry'] = Industry.objects.filter(isdeleted=0).order_by('description')
        context['vat'] = Vat.objects.filter(isdeleted=0).order_by('pk')
        return context


@method_decorator(login_required, name='dispatch')
class UpdateView(UpdateView):
    model = Customer
    template_name = 'customer/edit.html'
    fields = ['code', 'name', 'address1', 'address2', 'address3', 'telno1',
              'telno2', 'telno3', 'faxno1', 'faxno2',
              'tin', 'pagerno', 'payterms', 'creditlimit',
              'creditstatus', 'creditrating', 'contactperson',
              'contactposition', 'contactemail', 'remarks',
              'beg_amount', 'beg_code', 'beg_date',
              'end_amount', 'end_code', 'end_date', 'bankaccount',
              'creditterm', 'currency', 'customertype', 'industry', 'vat']

    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('customer.change_customer'):
            raise Http404
        return super(UpdateView, self).dispatch(request, *args, **kwargs)

    def form_valid(self, form):
        self.object = form.save(commit=False)
        self.object.multiplestatus = 'Y'
        self.object.enterby = self.request.user
        self.object.modifyby = self.request.user

        if self.object.vat:
            self.object.vatrate = self.object.vat.rate

        self.object.save(update_fields=['name', 'address1', 'address2', 'address3', 'telno1',
                                        'telno2', 'telno3',
                                        'faxno1', 'faxno2', 'tin', 'pagerno',
                                        'payterms', 'creditlimit', 'creditstatus',
                                        'creditrating', 'contactperson', 'contactposition',
                                        'contactemail', 'remarks',
                                        'multiplestatus', 'beg_amount', 'beg_code',
                                        'beg_date', 'end_amount',
                                        'end_code', 'end_date', 'bankaccount',
                                        'creditterm', 'currency', 'customertype',
                                        'industry', 'modifyby', 'modifydate', 'vat'])
        return HttpResponseRedirect('/customer')

    def get_context_data(self, **kwargs):
        context = super(UpdateView, self).get_context_data(**kwargs)
        context['customertype'] = Customertype.objects.filter(isdeleted=0).order_by('description')
        context['creditterm'] = Creditterm.objects.filter(isdeleted=0).order_by('description')
        context['currency'] = Currency.objects.filter(isdeleted=0).order_by('description')
        context['bankaccount'] = Bankaccount.objects.filter(isdeleted=0).order_by('bank')
        context['industry'] = Industry.objects.filter(isdeleted=0).order_by('description')
        context['vat'] = Vat.objects.filter(isdeleted=0).order_by('pk')
        return context


@method_decorator(login_required, name='dispatch')
class DeleteView(DeleteView):
    model = Customer
    template_name = 'customer/delete.html'

    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('customer.delete_customer'):
            raise Http404
        return super(DeleteView, self).dispatch(request, *args, **kwargs)

    def delete(self, request, *args, **kwargs):
        self.object = self.get_object()
        self.object.modifyby = self.request.user
        self.object.modifydate = datetime.datetime.now()
        self.object.isdeleted = 1
        self.object.status = 'I'
        self.object.save()
        return HttpResponseRedirect('/customer')


@method_decorator(login_required, name='dispatch')
class GeneratePDF(View):
    def get(self, request):
        company = Companyparameter.objects.all().first()
        list = Customer.objects.filter(isdeleted=0).order_by('code')
        # list = Customer.objects.filter(isdeleted=0, code__lte='B').order_by('code')

        context = {
            "title": "Customer Master List",
            "today": timezone.now(),
            "company": company,
            "list": list,
            "username": request.user,
        }
        return Render.render('customer/list.html', context)
