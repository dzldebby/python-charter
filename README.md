# Customised Powerpoint Chart and Table Generation using Python
*Automated generation of customised stacked bar chart and table in Microsoft Powerpoint using Python-pptx* 

Python-pptx is a Python library for generating a customized PowerPoint presentation from a database, provided by [Steve Canny](https://python-pptx.readthedocs.io/en/latest/) This module uses Python-pptx to generate a stacked bar chart, as well as table containing the data on a powerpoint slide, based on a csv input by the user. 

The stacked bar chart format is most suitable for the reporting of a 5-point Likert scale in the traditional Top 2 Box, Neutral, Bottom 2 Box format. The table replicates data presented in the table format as-is. 


## Example 

Given `data.csv`

```
This is your auto-generated slide,,,
Scale,X,Y,Z
T2B,20,20,50
Neutral,10,20,5
B2B,70,60,45
```

Generate slide as below:
![alt text](images/beforeafter.png?raw=True "Powerpoint format.")

### Generating multiple slides with one csv file 

It is possible to generate multiple slides (each with its own stacked bar chart and table), with just a single csv file as input. 

To do this, you will need to: 

#### Separate your data that you want on each slide with 2 rows  

```
This is your auto-generated slide,,,
Scale,X,Y,Z
T2B,20,20,50
Neutral,10,20,5
B2B,70,60,45
,,,
,,,
2nd auto-generated slide,,,
Scale,A,B,C
T2B,20,20,20
Neutral,20,20,20
B2B,60,60,60
,,,
,,,
```



## Usage 

1. Install requirements 

```python
pip install requirements.txt
```

2. Powerpoint template 

It is important to use the powerpoint template 'chart-01.pptx' as the base powerpoint, because a table placeholder has been created to ensure that that position of the table. 


