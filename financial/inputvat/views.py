import datetime
from django.views.generic import ListView, DetailView, CreateView, UpdateView, DeleteView
from django.contrib.auth.decorators import login_required
from django.utils.decorators import method_decorator
from django.http import HttpResponseRedirect, Http404
from inputvat.models import Inputvat
from inputvattype.models import Inputvattype
from chartofaccount.models import Chartofaccount

@method_decorator(login_required, name='dispatch')
class IndexView(ListView):
    model = Inputvat
    template_name = 'inputvat/index.html'
    context_object_name = 'data_list'

    def get_queryset(self):
        return Inputvat.objects.all().filter(isdeleted=0).order_by('-pk')


@method_decorator(login_required, name='dispatch')
class DetailView(DetailView):
    model = Inputvat
    template_name = 'inputvat/detail.html'


@method_decorator(login_required, name='dispatch')
class CreateView(CreateView):
    model = Inputvat
    template_name = 'inputvat/create.html'
    fields = ['code', 'description', 'inputvattype', 'inputvatchartofaccount', 'title']

    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('inputvat.add_inputvat'):
            raise Http404
        return super(CreateView, self).dispatch(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        context = super(CreateView, self).get_context_data(**kwargs)
        context['inputvattype'] = Inputvattype.objects.\
            filter(isdeleted=0).order_by('description')
        if self.request.POST.get('inputvatchartofaccount', False):
            context['inputvatchartofaccount'] = Chartofaccount.objects.get(pk=self.request.POST['inputvatchartofaccount'], isdeleted=0)
        return context

    def form_valid(self, form):
        self.object = form.save(commit=False)
        self.object.enterby = self.request.user
        self.object.modifyby = self.request.user
        self.object.save()
        return HttpResponseRedirect('/inputvat')


@method_decorator(login_required, name='dispatch')
class UpdateView(UpdateView):
    model = Inputvat
    template_name = 'inputvat/edit.html'
    fields = ['code', 'description', 'inputvattype', 'inputvatchartofaccount', 'title']

    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('inputvat.change_inputvat'):
            raise Http404
        return super(UpdateView, self).dispatch(request, *args, **kwargs)

    def get_context_data(self, **kwargs):
        context = super(UpdateView, self).get_context_data(**kwargs)
        context['inputvattype'] = Inputvattype.objects.\
            filter(isdeleted=0).order_by('description')
        if self.request.POST.get('inputvatchartofaccount', False):
            context['inputvatchartofaccount'] = Chartofaccount.objects.get(pk=self.request.POST['inputvatchartofaccount'], isdeleted=0)
        elif self.object.inputvatchartofaccount:
            context['inputvatchartofaccount'] = Chartofaccount.objects.get(pk=self.object.inputvatchartofaccount.id, isdeleted=0)
        return context

    def form_valid(self, form):
        self.object = form.save(commit=False)
        self.object.modifyby = self.request.user
        self.object.modifydate = datetime.datetime.now()
        self.object.save(update_fields=['description', 'inputvattype', 'inputvatchartofaccount', 
                                        'title', 'modifyby',
                                        'modifydate'])
        return HttpResponseRedirect('/inputvat')


@method_decorator(login_required, name='dispatch')
class DeleteView(DeleteView):
    model = Inputvat
    template_name = 'inputvat/delete.html'

    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('inputvat.delete_inputvat'):
            raise Http404
        return super(DeleteView, self).dispatch(request, *args, **kwargs)

    def delete(self, request, *args, **kwargs):
        self.object = self.get_object()
        self.object.modifyby = self.request.user
        self.object.modifydate = datetime.datetime.now()
        self.object.isdeleted = 1
        self.object.status = 'I'
        self.object.save()
        return HttpResponseRedirect('/inputvat')
