import datetime
from django.views.generic import ListView, DetailView, CreateView, UpdateView, DeleteView
from django.contrib.auth.decorators import login_required
from django.utils.decorators import method_decorator
from django.http import HttpResponseRedirect, Http404
from . models import Apsubtype


# Create your views here.
@method_decorator(login_required, name='dispatch')
class IndexView(ListView):
    model = Apsubtype
    template_name = 'apsubtype/index.html'
    context_object_name = 'data_list'

    def get_queryset(self):
        return Apsubtype.objects.all().filter(isdeleted=0).order_by('-pk')


@method_decorator(login_required, name='dispatch')
class DetailView(DetailView):
    model = Apsubtype
    template_name = 'apsubtype/detail.html'


@method_decorator(login_required, name='dispatch')
class CreateView(CreateView):
    model = Apsubtype
    template_name = 'apsubtype/create.html'
    fields = ['code', 'description']

    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('apsubtype.add_apsubtype'):
            raise Http404
        return super(CreateView, self).dispatch(request, *args, **kwargs)

    def form_valid(self, form):
        self.object = form.save(commit=False)
        self.object.enterby = self.request.user
        self.object.modifyby = self.request.user
        self.object.save()
        return HttpResponseRedirect('/apsubtype')


@method_decorator(login_required, name='dispatch')
class UpdateView(UpdateView):
    model = Apsubtype
    template_name = 'apsubtype/edit.html'
    fields = ['code', 'description']

    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('apsubtype.change_apsubtype'):
            raise Http404
        return super(UpdateView, self).dispatch(request, *args, **kwargs)

    def form_valid(self, form):
        self.object = form.save(commit=False)
        self.object.modifyby = self.request.user
        self.object.modifydate = datetime.datetime.now()
        self.object.save(update_fields=['description', 'modifyby', 'modifydate'])
        return HttpResponseRedirect('/apsubtype')


@method_decorator(login_required, name='dispatch')
class DeleteView(DeleteView):
    model = Apsubtype
    template_name = 'apsubtype/delete.html'

    def dispatch(self, request, *args, **kwargs):
        if not request.user.has_perm('apsubtype.delete_apsubtype'):
            raise Http404
        return super(DeleteView, self).dispatch(request, *args, **kwargs)

    def delete(self, request, *args, **kwargs):
        self.object = self.get_object()
        self.object.modifyby = self.request.user
        self.object.modifydate = datetime.datetime.now()
        self.object.isdeleted = 1
        self.object.status = 'I'
        self.object.save()
        return HttpResponseRedirect('/apsubtype')
