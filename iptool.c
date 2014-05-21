#include <stdlib.h>
#include <stdio.h>

int main(void)
{
    FILE    *fpin, *fpout;
    unsigned start, end;
    unsigned ip[4];
    int      n;
    char     code[16];
    char     line[256];

    fpin  = fopen("ip.txt", "rb");
    fpout = fopen("ip.out.txt", "wb");
    if (!fpin || !fpout) goto done;

    while (!feof(fpin)) {
        fscanf(fpin, "%s %d.%d.%d.%d %d", &code, &(ip[0]), &(ip[1]), &(ip[2]), &(ip[3]), &n);
        fgets(line, 256, fpin);
//      fprintf(fpout, "%s %d.%d.%d.%d %d\r\n", code, ip[0], ip[1], ip[2], ip[3], n);
        start = (ip[0] << 24) | (ip[1] << 16) | (ip[2] << 8) | (ip[3] << 0);
        end   = start + n;
        fprintf(fpout, "%u\t%u\t%d.%d.%d.%d\t%d.%d.%d.%d\t%s\r\n",
            start, end,
            (start >> 24) & 0xff, (start >> 16) & 0xff, (start >> 8) & 0xff, start & 0xff,
            (end >> 24) & 0xff, (end >> 16) & 0xff, (end >> 8) & 0xff, end & 0xff,
            code);
    }

done:
    if (fpin ) fclose(fpin );
    if (fpout) fclose(fpout);
}
