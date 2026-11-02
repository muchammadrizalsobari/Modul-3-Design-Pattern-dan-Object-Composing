import java.util.Scanner;

public class KasirAwal {
    public static void main(String[] args) {
        String nama;
        int harga;
        int stok;
        Scanner in = new Scanner(System.in);

        // Inisiasi
        nama = "GeForce GTX 1060Ti";
        harga = 3000;
        stok = 10;

        // Tampilan info
        System.out.println("Nama Barang : " + nama);
        System.out.println("Harga Barang : " + harga);
        System.out.println("Stok Barang : " + stok);

        // Transaksi pembelian
        int qty;
        System.out.println("Transaksi Pembelian");
        System.out.print("Jumlah barang : ");
        qty = in.nextInt();
        stok = stok - qty;

        // Print Pembayaran
        System.out.println("Jumlah Bayar : " + qty * harga);

        // Tampilan info
        System.out.println("Nama Barang : " + nama);
        System.out.println("Harga Barang : " + harga);
        System.out.println("Stok Barang : " + stok);
    }
}
