%penyelesaian kasus menggunakan metode WP
n1 = xlsread('Data_WP.xlsx','Sheet1','C:E'); %mengambil data dari xlsx tabel C, D, dan E
n2 = xlsread('Data_WP.xlsx','Sheet1','H:H'); %mengambil data dari xlsx tabel H
x = cat(2,n1,n2); %berfungsi untuk menggabungkan tabel C,D,E dengan tabel H
x(51:414,:) = [] %berfungsi untuk mengambil data dari 1 - 50
k = [0,0,1,0]; %pengaturan atribut dimana 1 merupakan benefit dan 0 adalah cost
w = [3,5,4,1]; %mengatur nilai bobot tiap kriteria 


[m n]=size (x); %matriks m x n dengan ukuran sebanyak variabel x (input)
w=w./sum(w); 

%melakukan perhitungan vektor(S) per baris (alternatif)
for j=1:n
    if k(j)==0 
        w(j)=-1*w(j);
    end
end
for i=1:m
    S(i) = prod(x(i,:).^w);
    if isinf(S(i)) % agar tidak ada nilai infinite
        S(i) = 0;
    end
end

V = S/sum(S);

% mencari nilai tertinggi
[nilai, index] = sort(V, 'descend');

disp("max value :"); %menampilkan nilai tertinggi 1 - 5
disp(nilai(1:5));

disp("index :");
disp(index(1:5)); %menampilkan index tertinggi 1 - 5