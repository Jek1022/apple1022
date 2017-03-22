# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2017-03-22 15:02
from __future__ import unicode_literals

import datetime
from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    dependencies = [
        ('customer', '0003_auto_20170303_1514'),
    ]

    operations = [
        migrations.AlterField(
            model_name='customer',
            name='creditterm',
            field=models.ForeignKey(blank=True, default=2, null=True, on_delete=django.db.models.deletion.CASCADE, related_name='creditterm_id', to='creditterm.Creditterm'),
        ),
        migrations.AlterField(
            model_name='customer',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 3, 22, 15, 2, 56, 936000)),
        ),
    ]
