# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2017-03-21 18:28
from __future__ import unicode_literals

import datetime
from django.conf import settings
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('requisitionform', '0005_auto_20170321_1553'),
    ]

    operations = [
        migrations.AlterField(
            model_name='rfdetail',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 3, 21, 18, 28, 28, 673000)),
        ),
        migrations.AlterField(
            model_name='rfdetail',
            name='postdate',
            field=models.DateTimeField(default=datetime.datetime(2017, 3, 21, 18, 28, 28, 673000)),
        ),
        migrations.AlterField(
            model_name='rfdetailtemp',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 3, 21, 18, 28, 28, 675000)),
        ),
        migrations.AlterField(
            model_name='rfmain',
            name='actualapprover',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='actual_approver', to=settings.AUTH_USER_MODEL),
        ),
        migrations.AlterField(
            model_name='rfmain',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 3, 21, 18, 28, 28, 670000)),
        ),
        migrations.AlterField(
            model_name='rfmain',
            name='postby',
            field=models.ForeignKey(blank=True, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='rfmain_post', to=settings.AUTH_USER_MODEL),
        ),
        migrations.AlterField(
            model_name='rfmain',
            name='postdate',
            field=models.DateTimeField(blank=True, null=True),
        ),
    ]
