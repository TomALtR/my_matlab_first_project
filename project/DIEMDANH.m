function varargout = DIEMDANH(varargin)
% DIEMDANH MATLAB code for DIEMDANH.fig
%      DIEMDANH, by itself, creates a new DIEMDANH or raises the existing
%      singleton*.
%
%      H = DIEMDANH returns the handle to a new DIEMDANH or the handle to
%      the existing singleton*.
%
%      DIEMDANH('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in DIEMDANH.M with the given input arguments.
%
%      DIEMDANH('Property','Value',...) creates a new DIEMDANH or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before DIEMDANH_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to DIEMDANH_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help DIEMDANH

% Last Modified by GUIDE v2.5 18-Oct-2022 15:21:30

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @DIEMDANH_OpeningFcn, ...
                   'gui_OutputFcn',  @DIEMDANH_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before DIEMDANH is made visible.
function DIEMDANH_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to DIEMDANH (see VARARGIN)

% Choose default command line output for DIEMDANH
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes DIEMDANH wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = DIEMDANH_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
% hObject    handle to pushbutton1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in exit.
function exit_Callback(hObject, eventdata, handles)
import java.util.*
close all
% hObject    handle to exit (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in select.
function select_Callback(hObject, eventdata, handles)
a= get(handles.ID,'value')
N= get(handles.name1,'string')
x1=get(handles.from,'string')
x1=str2double(x1)
x2=get(handles.to,'string')
x2=str2double(x2)
if a==1
    chon=randi([x1,x2],1)
    A=xlsread('DSSV.xls','A5:A133')
    set(handles.I_D,'string',num2str(A(chon)))
    [txt,B]=xlsread('DSSV.xls','C5:C133');
    set(handles.name2,'string',B(chon))
    c=mod(A(chon),5)
    if c==0 
        set(handles.team,'string',' 1')
    elseif c==1 
        set(handles.team,'string',' 2')
          elseif c==2 
        set(handles.team,'string',' 3')
          elseif c==3 
        set(handles.team,'string',' 4')
    else c==4 
        set(handles.team,'string',' 5')
    end
    
else
Key  = N(randi(numel(N), 1, 1))
  [txt,B]=xlsread('DSSV.xls','C5:C133');
    set(handles.name2,'string',B(Key))
 A=xlsread('DSSV.xls','A5:A133')
    set(handles.I_D,'string',num2str(A(Key)))
      c=mod(A(N),5)
    if c==0 
        set(handles.team,'string',' 1')
    elseif c==1 
        set(handles.team,'string',' 2')
         elseif c==2 
        set(handles.team,'string',' 3')
         elseif c==3 
        set(handles.team,'string',' 4')
    else c==4 
        set(handles.team,'string',' 5')
    end
end


    
% hObject    handle to select (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on button press in name.
function name_Callback(hObject, eventdata, handles)
set (handles.ID,'value',0)
% hObject    handle to name (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of name


% --- Executes on button press in ID.
function ID_Callback(hObject, eventdata, handles)
set(handles.name,'value',0)
% hObject    handle to ID (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of ID



function name2_Callback(hObject, eventdata, handles)

% hObject    handle to name2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of name2 as text
%        str2double(get(hObject,'String')) returns contents of name2 as a double


% --- Executes during object creation, after setting all properties.
function name2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to name2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function name1_Callback(hObject, eventdata, handles)
a=get(handles.name1,'string')
% hObject    handle to name1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of name1 as text
%        str2double(get(hObject,'String')) returns contents of name1 as a double


% --- Executes during object creation, after setting all properties.
function name1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to name1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function from_Callback(hObject, eventdata, handles)
a=get(handles.from,'string')
a=str2double(a)
set(handles.slider1,'value',a)
% hObject    handle to from (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of from as text
%        str2double(get(hObject,'String')) returns contents of from as a double


% --- Executes during object creation, after setting all properties.
function from_CreateFcn(hObject, eventdata, handles)
% hObject    handle to from (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function to_Callback(hObject, eventdata, handles)
a=get(handles.to,'string')
a=str2double(a)
set(handles.slider2,'value',a)
% hObject    handle to to (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of to as text
%        str2double(get(hObject,'String')) returns contents of to as a double


% --- Executes during object creation, after setting all properties.
function to_CreateFcn(hObject, eventdata, handles)
% hObject    handle to to (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in check.
function check_Callback(hObject, eventdata, handles)

% hObject    handle to check (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hint: get(hObject,'Value') returns toggle state of check



function mark_Callback(hObject, eventdata, handles)
a=get(handles.mark,'string')
a= str2double(a)
% hObject    handle to mark (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of mark as text
%        str2double(get(hObject,'String')) returns contents of mark as a double


% --- Executes during object creation, after setting all properties.
function mark_CreateFcn(hObject, eventdata, handles)
% hObject    handle to mark (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function team_Callback(hObject, eventdata, handles)
% hObject    handle to team (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of team as text
%        str2double(get(hObject,'String')) returns contents of team as a double


% --- Executes during object creation, after setting all properties.
function team_CreateFcn(hObject, eventdata, handles)
% hObject    handle to team (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function I_D_Callback(hObject, eventdata, handles)
% hObject    handle to I_D (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of I_D as text
%        str2double(get(hObject,'String')) returns contents of I_D as a double


% --- Executes during object creation, after setting all properties.
function I_D_CreateFcn(hObject, eventdata, handles)
% hObject    handle to I_D (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in confirm.
function confirm_Callback(hObject, eventdata, handles)
a=get(handles.check,'value')
b=get(handles.I_D,'string')
b=str2double(b)
c=get(handles.name2,'string')
A=[b,c,1]
if a==1
    xlswrite('DSSV1.xls',A,'attendance','A2')    
end
% hObject    handle to confirm (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)


% --- Executes on slider movement.
function slider1_Callback(hObject, eventdata, handles)
a=get(handles.slider1,'value')
a=num2str(a)
set(handles.from,'string',a)
% hObject    handle to slider1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'Value') returns position of slider
%        get(hObject,'Min') and get(hObject,'Max') to determine range of slider


% --- Executes during object creation, after setting all properties.
function slider1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to slider1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: slider controls usually have a light gray background.
if isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor',[.9 .9 .9]);
end
% --- Executes on slider movement.
function slider2_Callback(hObject, eventdata, handles)
a=get(handles.slider2,'value')
a=num2str(a)
set(handles.to,'string',a)
% hObject    handle to slider2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'Value') returns position of slider
%        get(hObject,'Min') and get(hObject,'Max') to determine range of slider


% --- Executes during object creation, after setting all properties.
function slider2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to slider2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: slider controls usually have a light gray background.
if isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor',[.9 .9 .9]);
end


% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)
% hObject    handle to listbox1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns listbox1 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from listbox1


% --- Executes during object creation, after setting all properties.
function listbox1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to listbox1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: listbox controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in confirm2.
function confirm2_Callback(hObject, eventdata, handles)
a=get(handles.mark,'string')
b=get(handles.I_D,'string')
b=str2double(b)
c=get(handles.name2,'string')
A=[b,c,1]
    A=[b,c,a]
    xlswrite('DSSV1.xls',A,'Mark','A2')


% hObject    handle to confirm2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
