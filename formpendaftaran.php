<!DOCTYPE HTML>  
<html>
<head>
    <title>Formulir Peserta Didik</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.0-beta3/dist/css/bootstrap.min.css"
    rel="stylesheet"
    integrity="sha384-eOJMYsd53ii+scO/bJGFsiCZc+5NDVN2yr8+0RDqr0Ql0h+rP48ckxlpbzKgwra6"
    crossorigin="anonymous">   
    <style>
    .warning {color: #FF0000;}
    </style>
</head>
<body>
    <?php
    //membuat variabel untuk error
    $error_jp = "";
    $error_tgl_masuk = "";
    $error_nis = "";
    $error_no_peserta = "";
    $error_paud = "";
    $error_tk = "";
    $error_no_skhun = "";
    $error_no_ijazah = "";
    $error_hobi = "";
    $error_cita = "";
    $error_nama = "";
    $error_jk = "";
    $error_nisn = "";
    $error_nik = "";
    $error_tmpt_lahir = "";
    $error_tgl_lahir = "";
    $error_agama = "";
    $error_bk = "";
    $error_alamat_jln = "";
    $error_rt = "";
    $error_rw = "";
    $error_dsn = "";
    $error_ds = "";
    $error_kec = "";
    $error_pos = "";
    $error_tmpt_tinggal = "";
    $error_transport = "";
    $error_hp = "";
    $error_telp = "";
    $error_email = "";
    $error_kps = "";
    $error_no_kps = "";
    $error_kwn = "";
    $error_negara = "";

    //membuat variabel data diri peserta didik
    $jp = "";
    $tgl_masuk = "";
    $nis = "";
    $no_peserta = "";
    $paud = "";
    $tk = "";
    $no_skhun = "";
    $no_ijazah = "";
    $hobi = "";
    $cita = "";
    $nama = "";
    $jk = "";
    $nisn = "";
    $nik = "";
    $tmpt_lahir = "";
    $tgl_lahir = "";
    $agama = "";
    $bk = "";
    $alamat_jln = "";
    $rt = "";
    $rw = "";
    $dsn = "";
    $ds = "";
    $kec = "";
    $pos = "";
    $tmpt_tinggal = "";
    $transport = "";
    $hp = "";
    $telp = "";
    $email = "";
    $kps = "";
    $no_kps = "";
    $kwn = "";
    $negara = "";

    if ($_SERVER["REQUEST_METHOD"] == "POST") {
        //jenis pendaftaran
        if (empty($_POST["jp"])) {
            $error_jp = "Jenis pendaftaran tidak boleh kosong";} 
        else{
            $jp = cek_input($_POST["jp"]);}
    
        //tanggal masuk sekolah
            if (empty($_POST["tgl_masuk"])) {
            $error_tgl_masuk = "Tanggal masuk sekolah tidak boleh kosong";} 
        else{
            $tgl_masuk = cek_input($_POST["tgl_masuk"]);
            if (!preg_match("/^[0-9+\/]/", $tgl_masuk)) {
            $error_tgl_masuk="Format tanggal tidak valid";}}
    
        //nis
        if (empty($_POST["nis"])) {
            $error_nis = "NIS tidak boleh kosong";} 
        else{
            $nis = cek_input($_POST["nis"]);
            if (!is_numeric($nis) || strlen($_POST["nis"])!=10) {
                $error_nis = "NIS hanya boleh angka dan berjumlah 10 digit";}}
    
        //nomor peserta ujian
        if (empty($_POST["no_peserta"])) {
            $error_no_peserta = "Nomor peserta ujian tidak boleh kosong";} 
        else{
            $no_peserta = cek_input($_POST["no_peserta"]);
            if (!is_numeric($no_peserta) || strlen($_POST["no_peserta"])!=20) {
                $error_no_peserta = "Nomor Peserta Ujian hanya boleh angka dan berjumlah 20 digit";}}
    
        //paud
        if (empty($_POST["paud"])) {
            $error_paud = "Apakah pernah PAUD tidak boleh kosong";} 
        else{
            $paud = cek_input($_POST["paud"]);}
    
        //tk
        if (empty($_POST["tk"])) {
            $error_tk = "Apakah pernah TK tidak boleh kosong";} 
        else{
            $tk = cek_input($_POST["tk"]);}
    
        //no skhun
        if (empty($_POST["no_skhun"])) {
            $error_no_skhun = "No. seri SKHUN sebelumnya tidak boleh kosong";} 
        else{
            $no_skhun = cek_input($_POST["no_skhun"]);
            if (!is_numeric($no_skhun) || strlen($_POST["no_skhun"])!=16) {
            $error_no_skhun = "No. seri SKHUN hanya boleh angka dan berjumlah 16 digit";}}
    
        //no ijazah
        if (empty($_POST["no_ijazah"])) {
            $error_no_ijazah = "No. seri ijazah sebelumnya tidak boleh kosong";} 
        else{
            $no_ijazah = cek_input($_POST["no_ijazah"]);
            if (!is_numeric($no_ijazah) || strlen($_POST["no_ijazah"])!=16) {
                $error_no_ijazah = "No. seri Ijazah hanya boleh angka dan berjumlah 16 digit";}}
    
        //hobi
        if (empty($_POST["hobi"])) {
            $error_hobi = "Hobi tidak boleh kosong";} 
        else{
            $hobi = cek_input($_POST["hobi"]);}
    
        //cita cita
        if (empty($_POST["cita"])) {
            $error_cita = "Cita-cita tidak boleh kosong";} 
        else{
            $cita = cek_input($_POST["cita"]);}
    
        //nama
        if (empty($_POST["nama"])) {
            $error_nama = "Nama tidak boleh kosong";} 
        else{
            $nama = cek_input($_POST["nama"]);
            if (!preg_match("/^[a-zA-Z]*$/",$nama)) {
                $error_nama = "Inputan hanya boleh huruf dan spasi";}}

        //jenis kelamin
        if (empty($_POST["jk"])) {
            $error_jk = "Jenis Kelamin tidak boleh kosong";} 
        else{
            $jk = cek_input($_POST["jk"]);}
    
        //nisn
        if (empty($_POST["nisn"])) {
            $error_nisn = "NISN tidak boleh kosong";} 
        else{
            $nisn = cek_input($_POST["nisn"]);
            if (!is_numeric($nisn) || strlen($_POST["nisn"])!=10) {
            $error_nisn = "NISN hanya boleh angka dan berjumlah 10 digit";}}
    
        //nik
        if (empty($_POST["nik"])) {
            $error_nik = "NIK tidak boleh kosong";} 
        else{
            $nik = cek_input($_POST["nik"]);
            if (!is_numeric($nik) || strlen($_POST["nik"])!=16) {
                $error_nik = "NIK hanya boleh angka dan berjumlah 16 digit";}}
    
        //tempat lahir
        if (empty($_POST["tmpt_lahir"])) {
            $error_tmpt_lahir = "Tempat lahir tidak boleh kosong";} 
        else{
            $tmpt_lahir = cek_input($_POST["tmpt_lahir"]);
            if (!preg_match("/^[a-zA-Z]*$/", $tmpt_lahir)) {
            $error_tmpt_lahir="Tempat lahir hanya boleh huruf dan spasi";}}
    
        //tanggal lahir
        if (empty($_POST["tgl_lahir"])) {
            $error_tgl_lahir = "Tanggal lahir tidak boleh kosong";} 
        else{
            $tgl_lahir = cek_input($_POST["tgl_lahir"]);
            if (!preg_match("/^[0-9+\/]/", $tgl_lahir)) {
            $error_tgl_lahir="Format tanggal tidak valid";}}
    
        //agama
        if (empty($_POST["agama"])) {
            $error_agama = "Agama tidak boleh kosong";} 
        else{
            $agama = cek_input($_POST["agama"]);}
    
        //berkebutuhan khusus
        if (empty($_POST["bk"])) {
            $error_bk = "Berkebutuhan khusus tidak boleh kosong";} 
        else{
            $bk = cek_input($_POST["bk"]);}
    
        //jalan
        if (empty($_POST["alamat_jln"])) {
            $error_alamat_jln = "Alamat jalan tidak boleh kosong";} 
        else{
            $alamat_jln = cek_input($_POST["alamat_jln"]);}
    
        //rt
        if (empty($_POST["rt"])) {
            $error_rt = "RT tidak boleh kosong";} 
        else{
            $rt = cek_input($_POST["rt"]);
            if(!is_numeric($rt)){
            $error_rt = "RT hanya boleh angka";}}
    
        //rw
        if (empty($_POST["rw"])) {
            $error_rw = "RW tidak boleh kosong";} 
        else{
            $rw = cek_input($_POST["rw"]);
            if(!is_numeric($rw)){
            $error_rw = "RW hanya boleh angka";}}
    
        //dusun
        if (empty($_POST["dsn"])) {
            $error_dsn = "Dusun tidak boleh kosong";} 
        else{
            $dsn = cek_input($_POST["dsn"]);}
    
        //desa
        if (empty($_POST["ds"])) {
            $error_ds = "Desa tidak boleh kosong";} 
        else{
            $ds = cek_input($_POST["ds"]);}
    
        //kecamatan
        if (empty($_POST["kec"])) {
            $error_kec = "Kecamatan tidak boleh kosong";} 
        else{
            $kec = cek_input($_POST["kec"]);}

        //kode pos
        if (empty($_POST["pos"])) {
            $error_pos = "Pos tidak boleh kosong";} 
        else{
            $pos = cek_input($_POST["pos"]);
            if (!is_numeric($pos) || strlen($_POST["pos"])!=5) {
            $error_pos = "Kode Pos hanya boleh angka dan berjumlah 5 digit";}}

        //tempat tinggal
        if (empty($_POST["tmpt_tinggal"])) {
            $error_tmpt_tinggal = "Tempat tinggal tidak boleh kosong";} 
        else{
            $tmpt_tinggal = cek_input($_POST["tmpt_tinggal"]);}

        //moda transportasi
        if (empty($_POST["transport"])) {
            $error_transport = "Moda transportasi tidak boleh kosong";} 
        else{
            $transport = cek_input($_POST["transport"]);}
    
        //no hp
        if (empty($_POST["hp"])) {
            $error_hp = "Nomor HP tidak boleh kosong";} 
        else{
            $hp = cek_input($_POST["hp"]);
            if (!is_numeric($hp)) {
            $error_hp = 'Nomor HP hanya boleh angka';}}
    
        //telepon
        if (empty($_POST["telp"])) {
            $error_telp = "Nomor telepon tidak boleh kosong";} 
        else{
            $telp = cek_input($_POST["telp"]);
            if (!is_numeric($telp)) {
            $error_telp = 'Nomor Telepon hanya boleh angka';}}
    
        //email
        if (empty($_POST["email"])) {
            $error_email = "Email tidak boleh kosong";} 
        else{
            $email = cek_input($_POST["email"]);}
    
        //penerima kps
        if (empty($_POST["kps"])) {
            $error_kps = "Penerima KPS/PKH/KIP tidak boleh kosong";}
        else {
            $kps = cek_input($_POST["kps"]);}
        
        //no kps
        if (($_POST["no_kps"])=="Y"&&empty($_POST["no_kps"])) {
            $error_no_kps="Pemilik KPS/PKH/KIP wajib mengisi";}
        elseif (($_POST["no_kps"])=="T") {
            $no_kps="-";}
        else {
            $no_kps = cek_input($_POST["no_kps"]);
            if (!preg_match("/^[a-z0-9]*$/", $no_kps)) {
            $error_no_kps="Format no KPS invalid";}}
    
        //kewarganegaraan
        if (empty($_POST["kwn"])) {
            $error_kwn = "Kewarganegaraan tidak boleh kosong";} 
        else{
            $kwn = cek_input($_POST["kwn"]);}

        //negara
        if ($kwn == "Asing") {
            if(empty($_POST["negara"])){
                $error_negara="Nama Negara tidak boleh kosong";}
            else{
                $negara=cek_input($_POST["negara"]);
                if(!preg_match("/^[a-zA-z ]*$/", $negara)){
                    $error_negara="Inputan Hanya Boleh Huruf dan Spasi";}}
        }
        elseif ($kwn == "Indo") {
            $negara="Indonesia";}
        }
    
        include 'koneksidatasiswa.php'; //menghubungkan dengan file koneksidatasiswa.php
        if (!empty($jp) && !empty($tgl_masuk) && !empty($nis) && !empty($no_peserta) && !empty($paud) && !empty($tk) &&
            !empty($no_skhun) && !empty($no_ijazah) && !empty($hobi) && !empty($cita) && !empty($nama) && !empty($jk) &&
            !empty($nisn) && !empty($nik) && !empty($tmpt_lahir) && !empty($tgl_lahir) && !empty($agama) && !empty($bk) &&
            !empty($alamat_jln) && !empty($rt) && !empty($rw) && !empty($dsn) && !empty($ds) && !empty($kec) &&
            !empty($pos) && !empty($tmpt_tinggal) && !empty($transport) && !empty($hp) && !empty($telp) && !empty($email) && 
            !empty($kps) && !empty($kwn)) {
            
            //query sql untuk menimpan data ke database
            $sql = "INSERT INTO data_peserta_didik set jenis_pendaftaran='$jp', tanggal_masuk='$tgl_masuk', nis='$nis', no_peserta='$no_peserta',
                paud='$paud', tk='$tk', no_skhun='$no_skhun', no_ijazah='$no_ijazah', hobi='$hobi', cita_cita='$cita', nama='$nama',
                jenis_kelamin='$jk', nisn='$nisn', nik='$nik', tempat_lahir='$tmpt_lahir', tanggal_lahir='$tgl_lahir', agama='$agama',
                berkebutuhan_khusus='$bk', alamat='$alamat_jln', rt='$rt', rw='$rw', nama_dusun='$dsn', kelurahan='$ds',kecamatan='$kec',
                kodepos='$pos', tempat_tinggal='$tmpt_tinggal', transportasi='$transport', hp='$hp', telp='$telp', email='$email', 
                penerima_kps='$kps', no_kps='$no_kps', kewarganegaraan='$kwn', negara='$negara' ";

            if (mysqli_query($koneksi, $sql)) {
                echo "<script> alert('Data Berhasil Tersimpan') </script>";
                ?><meta http-equiv="refresh" content="5"; url=formpendaftaran.php><?php
                // jika sukses maka muncul pemberitahuan "Data Berhasil Tersimpan"
            }
            else {
                echo "<script> alert('Data Gagal disimpan') </script>";
                // jika gagal maka muncul pemberitahuan "Data Gagal disimpan"}
        }
    }

    //membuat function cek input data
    function cek_input($data) {
        $data = trim($data);
        $data = stripslashes($data);
        $data = htmlspecialchars($data);
        return $data;
      }
    ?>

    <div class="row">
    <div class="col-md-13">
    <div class="card">
    <div class="card-header" align="center">Formulir Peserta Didik</div>
    <div class="card-body">
        <!-- membuat form dengan method post untuk membuat inputan data -->
        <form method="post" action="<?php echo htmlspecialchars($_SERVER["PHP_SELF"]);?>">  
            <div class="form-group row">
                <label for="jp" class="col-sm-2 col-form-label">Jenis Pendaftaran</label>
                <div class="col-sm-10">
                    <input type="radio" name="jp" <?php if (isset($jp) && $jp=="baru") ;?> value="Siswa Baru">Siswa Baru
                    <input type="radio" name="jp" <?php if (isset($jp) && $jp=="pindah") ;?> value="Pindahan">Pindahan 
                    <span class="warning"><?php echo $error_jp; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="tgl_masuk" class="col-sm-2 col-form-label">Tanggal masuk sekolah</label>
                <div class="col-sm-10">
                    <input type="text" name="tgl_masuk" class="form-control <?php echo ($error_tgl_masuk !="" ? "is-invalid" : ""); ?>" id="tgl_masuk" placeholder="Tanggal-Bulan-Tahun" value="<?php echo $tgl_masuk; ?>"><span class="warning"><?php echo $error_tgl_masuk; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="nis" class="col-sm-2 col-form-label">NIS</label>
                <div class="col-sm-10">
                    <input type="text" name="nis" class="form-control <?php echo ($error_nis !="" ? "is-invalid" : ""); ?>" id="nis" placeholder="NIS hanya boleh angka dan berjumlah 10 digit" value="<?php echo $nis; ?>"><span class="warning"><?php echo $error_nis; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="no_peserta" class="col-sm-2 col-form-label">Nomor Peserta Ujian</label>
                <div class="col-sm-10">
                    <input type="text" name="no_peserta" class="form-control <?php echo ($error_no_peserta !="" ? "is-invalid" : ""); ?>" id="no_peserta" placeholder="Nomor peserta ujian adalah 20 digit yang tertera dalam sertifikat SKHUN SD, diisi bagi peserta didik jenjang SMP" value="<?php echo $no_peserta; ?>"><span class="warning"><?php echo $error_no_peserta; ?></span>
                </div>
                
            </div>

            <div class="form-group row">
                <label for="paud" class="col-sm-2 col-form-label">Apakah pernah PAUD</label>
                <div class="col-sm-10">
                    <input type="radio" name="paud" <?php if (isset($paud) && $paud=="Y") ;?> value="Ya"> Ya
                    <input type="radio" name="paud" <?php if (isset($paud) && $paud=="T") ;?> value="Tidak"> Tidak 
                    <span class="warning"><?php echo $error_paud; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="tk" class="col-sm-2 col-form-label">Apakah pernah TK</label>
                <div class="col-sm-10">
                    <input type="radio" name="tk" <?php if (isset($tk) && $tk=="Y") ;?> value="Ya"> Ya
                    <input type="radio" name="tk" <?php if (isset($tk) && $tk=="T") ;?> value="Tidak"> Tidak 
                    <span class="warning"><?php echo $error_tk; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="no_skhun" class="col-sm-2 col-form-label">No. seri SKHUN sebelumnya</label>
                <div class="col-sm-10">
                    <input type="text" name="no_skhun" class="form-control <?php echo ($error_no_skhun !="" ? "is-invalid" : ""); ?>" id="no_skhun" placeholder="Diisi 16 digit yang tertera di SKHUN SD - bagi peserta didik jenjang SMP" value="<?php echo $no_skhun; ?>"><span class="warning"><?php echo $error_no_skhun; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="no_ijazah" class="col-sm-2 col-form-label">No. seri ijazah sebelumnya</label>
                <div class="col-sm-10">
                    <input type="text" name="no_ijazah" class="form-control <?php echo ($error_no_ijazah !="" ? "is-invalid" : ""); ?>" id="no_ijazah" placeholder="Diisi 16 digit yang tertera di ijazah SD - bagi peserta didik jenjang SMP" value="<?php echo $no_ijazah; ?>"><span class="warning"><?php echo $error_no_ijazah; ?></span>
                </div>
            </div>

            <div class="form-group row">
            <label for="hobi" class="col-sm-2 col-form-label">Hobi</label>
                <div class="col-sm-10">
                    <select name="hobi" class="form-control <?php echo ($error_hobi !="" ? "is-invalid":"");?>" id="hobi">
                        <option value="Olahraga"> Olahraga </option>
                        <option value="Kesenian"> Kesenian </option>
                        <option value="Membaca"> Membaca </option>
                        <option value="Menulis"> Menulis </option>
                        <option value="Traveling"> Traveling </option>
                        <option value="Lainnya"> Lainnya </option>
                    </select>
                    <span class="warning"><?php echo $error_hobi; ?></span>
                </div>
            </div>

            <div class="form-group row">
            <label for="cita" class="col-sm-2 col-form-label">Cita-Cita</label>
                <div class="col-sm-10">
                    <select name="cita" class="form-control <?php echo ($error_cita !="" ? "is-invalid":"");?>" id="cita">
                        <option value="PNS"> PNS </option>
                        <option value="TNI"> TNI/POLRI </option>
                        <option value="Guru"> Guru/Dosen </option>
                        <option value="Dokter"> Dokter </option>
                        <option value="Politikus"> Politikus </option>
                        <option value="Wiraswasta"> Wiraswasta </option>
                        <option value="Seni"> Seni/Lukis/Artis/Sejenis </option>
                        <option value="Lainnya"> Lainnya </option>
                    </select>
                    <span class="warning"><?php echo $error_cita; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="nama" class="col-sm-2 col-form-label">Nama</label>
                <div class="col-sm-10">
                    <input type="text" name="nama" class="form-control <?php echo ($error_nama !="" ? "is-invalid" : ""); ?>" id="nama" placeholder="Nama Lengkap" value="<?php echo $nama; ?>"><span class="warning"><?php echo $error_nama; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="jk" class="col-sm-2 col-form-label"> Jenis Kelamin </label>
                <div class="col-sm-10">
                    <input type="radio" name="jk" <?php if (isset($jk) && $jk=="L") ;?> value="Laki - laki"> Laki-laki
                    <input type="radio" name="jk" <?php if (isset($jk) && $jk=="P") ;?> value="Perempuan"> Perempuan 
                    <span class="warning"><?php echo $error_jk; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="nisn" class="col-sm-2 col-form-label">NISN</label>
                <div class="col-sm-10">
                    <input type="text" name="nisn" class="form-control <?php echo ($error_nisn !="" ? "is-invalid" : ""); ?>" id="nisn" placeholder="NISN hanya boleh angka dan berjumlah 10 digit" value="<?php echo $nisn; ?>"><span class="warning"><?php echo $error_nisn; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="nik" class="col-sm-2 col-form-label">NIK</label>
                <div class="col-sm-10">
                    <input type="text" name="nik" class="form-control <?php echo ($error_nik !="" ? "is-invalid" : ""); ?>" id="nik" placeholder="NIK hanya boleh angka dan berjumlah 16 digit" value="<?php echo $nik; ?>"><span class="warning"><?php echo $error_nik; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="tmpt_lahir" class="col-sm-2 col-form-label">Tempat Lahir</label>
                <div class="col-sm-10">
                    <input type="text" name="tmpt_lahir" class="form-control <?php echo ($error_tmpt_lahir !="" ? "is-invalid" : ""); ?>" id="tmpt_lahir" placeholder="Tempat lahir" value="<?php echo $tmpt_lahir; ?>"><span class="warning"><?php echo $error_tmpt_lahir; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="tgl_lahir" class="col-sm-2 col-form-label">Tanggal Lahir</label>
                <div class="col-sm-10">
                    <input type="text" name="tgl_lahir" class="form-control <?php echo ($error_tgl_lahir !="" ? "is-invalid" : ""); ?>" id="tgl_lahir" placeholder="Tanggal-Bulan-Tahun" value="<?php echo $tgl_lahir; ?>"><span class="warning"><?php echo $error_tgl_lahir; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="agama" class="col-sm-2 col-form-label"> Agama </label>
                <div class="col-sm-10">
                    <select name="agama" class="form-control <?php echo ($error_agama !="" ? "is-invalid":"");?>" id="agama">
                        <option value="Islam"> Islam </option>
                        <option value="Kristen"> Kristen / Protestan </option>
                        <option value="Katholik"> Katholik </option>
                        <option value="Hindu"> Hindu </option>
                        <option value="Budha"> Budha </option>
                        <option value="Khonghucu"> Khong Hu Cu </option>
                    </select>
                    <span class="warning"><?php echo $error_agama; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="bk" class="col-sm-2 col-form-label"> Berkebutuhan Khusus </label>
                <div class="col-sm-10">
                    <select name="bk" class="form-control <?php echo ($error_bk !="" ? "is-invalid":"");?>" id="bk">
                        <option value="Tidak"> Tidak </option>
                        <option value="Netra"> Netra </option>
                        <option value="Rungu"> Rungu </option>
                        <option value="Grahita Ringan"> Grahita Ringan </option>
                        <option value="Grahita Sedang"> Grahita Sedang </option>
                        <option value="Daksa Ringan"> Daksa Ringan </option>
                        <option value="Daksa Sedang"> Daksa Sedang </option>
                        <option value="Laras"> Laras </option>
                        <option value="Wicara"> Wicara </option>
                        <option value="Tuna Ganda"> Tuna Ganda </option>
                        <option value="Hiper Aktif"> Hiper Aktif </option>
                        <option value="Cerdas Istimewa"> Cerdas Istimewa </option>
                        <option value="Bakat Istemewa"> Bakat Istemewa </option>
                        <option value="Kesulitan Belajar"> Kesulitan Belajar </option>
                        <option value="Narkoba"> Narkoba </option>
                        <option value="Indigo"> Indigo </option>
                        <option value="Down Sindrome"> Down Sindrome </option>
                        <option value="Autis"> Autis </option>
                    </select>
                    <span class="warning"><?php echo $error_bk; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="alamat_jln" class="col-sm-2 col-form-label">Alamat Jalan</label>
                <div class="col-sm-10">
                    <input type="text" name="alamat_jln" class="form-control <?php echo ($error_alamat_jln !="" ? "is-invalid" : ""); ?>" id="alamat_jln" placeholder="Jalan" value="<?php echo $alamat_jln; ?>"><span class="warning"><?php echo $error_alamat_jln; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="rt" class="col-sm-2 col-form-label">RT</label>
                <div class="col-sm-10">
                    <input type="text" name="rt" class="form-control <?php echo ($error_rt !="" ? "is-invalid" : ""); ?>" id="rt" placeholder="RT" value="<?php echo $rt; ?>"><span class="warning"><?php echo $error_rt; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="rw" class="col-sm-2 col-form-label">RW</label>
                <div class="col-sm-10">
                    <input type="text" name="rw" class="form-control <?php echo ($error_rw !="" ? "is-invalid" : ""); ?>" id="rw" placeholder="RW" value="<?php echo $rw; ?>"><span class="warning"><?php echo $error_rw; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="dsn" class="col-sm-2 col-form-label">Dusun</label>
                <div class="col-sm-10">
                    <input type="text" name="dsn" class="form-control <?php echo ($error_dsn !="" ? "is-invalid" : ""); ?>" id="dsn" placeholder="Dusun" value="<?php echo $dsn; ?>"><span class="warning"><?php echo $error_dsn; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="ds" class="col-sm-2 col-form-label">Desa</label>
                <div class="col-sm-10">
                    <input type="text" name="ds" class="form-control <?php echo ($error_ds !="" ? "is-invalid" : ""); ?>" id="ds" placeholder="Desa" value="<?php echo $ds; ?>"><span class="warning"><?php echo $error_ds; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="kec" class="col-sm-2 col-form-label">Kecamatan</label>
                <div class="col-sm-10">
                    <input type="text" name="kec" class="form-control <?php echo ($error_kec !="" ? "is-invalid" : ""); ?>" id="kec" placeholder="Kecamatan" value="<?php echo $kec; ?>"><span class="warning"><?php echo $error_kec; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="pos" class="col-sm-2 col-form-label">Kode Pos</label>
                <div class="col-sm-10">
                    <input type="text" name="pos" class="form-control <?php echo ($error_pos !="" ? "is-invalid" : ""); ?>" id="pos" placeholder="Kode Pos" value="<?php echo $pos; ?>"><span class="warning"><?php echo $error_pos; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="tmpt_tinggal" class="col-sm-2 col-form-label"> Tempat Tinggal </label>
                <div class="col-sm-10">
                    <select name="tmpt_tinggal" class="form-control <?php echo ($error_tmpt_tinggal !="" ? "is-invalid":"");?>" id="tmpt_tinggal">
                        <option value="Bersama Orang Tua"> Bersama Orang Tua </option>
                        <option value="Wali"> Wali </option>
                        <option value="Kos"> Kos </option>
                        <option value="Asrama"> Asrama </option>
                        <option value="Panti Asuhan"> Panti Asuhan </option>
                        <option value="Lainnya"> Lainnya </option>
                    </select>
                    <span class="warning"><?php echo $error_tmpt_tinggal; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="transport" class="col-sm-2 col-form-label"> Moda Transportasi </label>
                <div class="col-sm-10">
                    <select name="transport" class="form-control <?php echo ($error_transport !="" ? "is-invalid":"");?>" id="transport">
                        <option value="Jalan"> Jalan Kaki </option>
                        <option value="Kendaraan Pribadi"> Kendaraan Pribadi </option>
                        <option value="Kendaraan Umum"> Kendaraan Umum/angkot/Pete-pete </option>
                        <option value="Jemputan"> Jemputan Sekolah </option>
                        <option value="Kereta"> Kereta Api </option>
                        <option value="Ojek"> Ojek </option>
                        <option value="Andong"> Andong/Bendi/Sado/Dokar/Delman/Becak </option>
                        <option value="Perahu"> Perahu Penyebrangan/Rakit/Getek </option>
                        <option value="Lainnya"> Lainnya </option>
                    </select>
                    <span class="warning"><?php echo $error_transport; ?></span>
                </div>
            </div>


            <div class="form-group row">
                <label for="hp" class="col-sm-2 col-form-label">Nomor HP</label>
                <div class="col-sm-10">
                    <input type="text" name="hp" class="form-control <?php echo ($error_hp !="" ? "is-invalid" : ""); ?>" id="hp" placeholder="Nomor HP" value="<?php echo $hp; ?>"><span class="warning"><?php echo $error_hp; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="telp" class="col-sm-2 col-form-label">Nomor Telepon</label>
                <div class="col-sm-10">
                    <input type="text" name="telp" class="form-control <?php echo ($error_telp !="" ? "is-invalid" : ""); ?>" id="telp" placeholder="No Telepon" value="<?php echo $telp; ?>"><span class="warning"><?php echo $error_telp; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="email" class="col-sm-2 col-form-label">Email Pribadi</label>
                <div class="col-sm-10">
                    <input type="text" name="email" class="form-control <?php echo ($error_email !="" ? "is-invalid" : ""); ?>" id="email" placeholder="Email" value="<?php echo $email; ?>"><span class="warning"><?php echo $error_email; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="kps" class="col-sm-2 col-form-label"> Penerima KPS/PKH/KIP </label>
                <div class="col-sm-10">
                    <input type="radio" name="kps" <?php if (isset($kps) && $kps=="Y") ;?> value="Ya"> Ya
                    <input type="radio" name="kps" <?php if (isset($kps) && $kps=="T") ;?> value="Tidak"> Tidak 
                    <span class="warning"><?php echo $error_kps; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="no_kps" class="col-sm-2 col-form-label">No. KPS/KKS/PKH/KIP</label>
                <div class="col-sm-10">
                    <input type="text" name="no_kps" class="form-control <?php echo ($error_no_kps !="" ? "is-invalid" : ""); ?>" id="no_kps" placeholder="Diisi apabila menerima KPS/KKS/PKH/KIP" value="<?php echo $no_kps; ?>"><span class="warning"><?php echo $error_no_kps; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="kwn" class="col-sm-2 col-form-label"> Kewarganegaraan </label>
                <div class="col-sm-10">
                    <input type="radio" name="kwn" <?php if (isset($kwn) && $kps=="Indo") ;?> value="Indonesia"> Indonesia
                    <input type="radio" name="kwn" <?php if (isset($kwn) && $kps=="Asing") ;?> value="Asing"> Asing 
                    <span class="warning"><?php echo $error_kwn; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <label for="negara" class="col-sm-2 col-form-label">Nama Negara</label>
                <div class="col-sm-10">
                    <input type="text" name="negara" class="form-control <?php echo ($error_negara !="" ? "is-invalid" : ""); ?>" id="negara" placeholder="Diisi apabila berkewarganegaraan asing" value="<?php echo $negara; ?>"><span class="warning"><?php echo $error_negara; ?></span>
                </div>
            </div>

            <div class="form-group row">
                <div class="col-sm-10">
                    <button type="submit" class="btn_kirim">KIRIM</button>
                </div>
            </div>
        </form>
    </div>
    </div>
    </div>
    </div>
    
    <?php  
    //Menampilkan hasil
    echo "<center><h2>Data yang Anda Masukkan:</h2></center>";
    echo "Jenis Pendaftaran = ". $jp;
    echo "<br>";
    echo "Tanggal Masuk Sekolah = ". $tgl_masuk;
    echo "<br>";
    echo "NIS = ". $nis;
    echo "<br>";
    echo "Nomer Peserta Ujian = ". $no_peserta;
    echo "<br>";
    echo "Pernah PAUD = ". $paud;
    echo "<br>";
    echo "Pernah TK = ". $tk;
    echo "<br>";
    echo "No. Seri SKHUN Sebelumnya = ". $no_skhun;
    echo "<br>";
    echo "No. Seri Ijazah Sebelumnya = ". $no_ijazah;
    echo "<br>";
    echo "Hobi = ". $hobi;
    echo "<br>";
    echo "Cita-cita = ", $cita;
    echo "<br>";
    echo "Nama Lengkap = ". $nama;
    echo "<br>";
    echo "Jenis Kelamin = ". $jk;
    echo "<br>";
    echo "NISN = ". $nisn;
    echo "<br>";
    echo "NIK = ". $nik;
    echo "<br>";
    echo "Tempat Lahir = ". $tmpt_lahir;
    echo "<br>";
    echo "Tanggal Lahir = ". $tgl_lahir;
    echo "<br>";
    echo "Agama = ". $agama;
    echo "<br>";
    echo "Berkebutuhan Khusus = ". $bk;
    echo "<br>";
    echo "Alamat Jalan = ". $alamat_jln;
    echo "<br>";
    echo "RT = ". $rt;
    echo "<br>";
    echo "RW = ". $rw;
    echo "<br>";
    echo "Dusun = ". $dsn;
    echo "<br>";
    echo "Desa/Kelurahan = ". $ds;
    echo "<br>";
    echo "Kecamatan = ". $kec;
    echo "<br>";
    echo "Kode Pos = ". $pos;
    echo "<br>";
    echo "Tempat Tinggal = ". $tmpt_tinggal;
    echo "<br>";
    echo "Transportasi = ". $transport;
    echo "<br>";
    echo "No. HP = ". $hp;
    echo "<br>";
    echo "E-mail Pribadi = ". $email;
    echo "<br>";
    echo "No. Telepon = ". $telp;
    echo "<br>";
    echo "Penerima KPS/PKH/KIP = ". $kps;
    echo "<br>";
    echo "Nomor KPS/KKS/PKH/KIP = ". $no_kps;
    echo "<br>";
    echo "Kewarganegaraan = ". $kwn;
    echo "<br>";
    echo "Nama Negara = ". $negara;
    ?>
    </br>
    <button onclick="location.href = 'reportdatasiswa.php';" id="export" class="btn-btn-primary">
    Export to Excel</button>
    </body>
</html>