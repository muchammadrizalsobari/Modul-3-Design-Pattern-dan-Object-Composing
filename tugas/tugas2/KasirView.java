public class KasirView {
    public void printBarangDetails(String nama, int harga, int stok) {
        System.out.println("Nama Barang : " + nama);
        System.out.println("Harga Barang : " + harga);
        System.out.println("Stok Barang : " + stok);
    }

    public void printTransaksiHeader() {
        System.out.println("Transaksi Pembelian");
    }

    public void printPembayaran(int totalBayar) {
        System.out.println("Jumlah Bayar : " + totalBayar);
    }

    public void printPesan(String pesan) {
        System.out.println(pesan);
    }
}
