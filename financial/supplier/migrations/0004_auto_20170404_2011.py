# -*- coding: utf-8 -*-
# Generated by Django 1.9.10 on 2017-04-04 20:11
from __future__ import unicode_literals

import datetime
from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('supplier', '0003_auto_20170303_1525'),
    ]

    operations = [
        migrations.AlterField(
            model_name='supplier',
            name='modifydate',
            field=models.DateTimeField(default=datetime.datetime(2017, 4, 4, 20, 11, 10, 731000)),
        ),
    ]
