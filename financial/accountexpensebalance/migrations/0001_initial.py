# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2018-05-17 03:04
from __future__ import unicode_literals

from django.conf import settings
import django.core.validators
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    initial = True

    dependencies = [
        ('chartofaccount', '0038_auto_20180323_1137'),
        migrations.swappable_dependency(settings.AUTH_USER_MODEL),
        ('department', '0034_auto_20180322_1106'),
    ]

    operations = [
        migrations.CreateModel(
            name='Accountexpensebalance',
            fields=[
                ('id', models.AutoField(auto_created=True, primary_key=True, serialize=False, verbose_name='ID')),
                ('year', models.PositiveSmallIntegerField(validators=[django.core.validators.MaxValueValidator(2100), django.core.validators.MinValueValidator(1980)])),
                ('month', models.PositiveIntegerField(validators=[django.core.validators.MaxValueValidator(1), django.core.validators.MinValueValidator(12)])),
                ('amount', models.DecimalField(blank=True, decimal_places=2, default=0.0, max_digits=18, null=True)),
                ('code', models.CharField(max_length=1)),
                ('date', models.DateField()),
                ('status', models.CharField(choices=[('A', 'Active'), ('I', 'Inactive'), ('C', 'Cancelled'), ('O', 'Posted'), ('P', 'Printed')], default='A', max_length=1)),
                ('enterdate', models.DateTimeField(auto_now_add=True)),
                ('modifydate', models.DateTimeField(auto_now_add=True)),
                ('isdeleted', models.IntegerField(default=0)),
                ('chartofaccount', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='accountexpensebalance_chartofaccount_id', to='chartofaccount.Chartofaccount')),
                ('department', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, related_name='accountexpensebalance_department_id', to='department.Department')),
                ('enterby', models.ForeignKey(default=1, on_delete=django.db.models.deletion.CASCADE, related_name='accountexpensebalance_enter', to=settings.AUTH_USER_MODEL)),
                ('modifyby', models.ForeignKey(default=1, on_delete=django.db.models.deletion.CASCADE, related_name='accountexpensebalance_modify', to=settings.AUTH_USER_MODEL)),
            ],
            options={
                'ordering': ['-pk'],
                'db_table': 'accountexpensebalance',
                'permissions': (('view_accountexpensebalance', 'Can view accountexpensebalance'),),
            },
        ),
        migrations.AlterUniqueTogether(
            name='accountexpensebalance',
            unique_together=set([('year', 'month', 'chartofaccount', 'department')]),
        ),
    ]
