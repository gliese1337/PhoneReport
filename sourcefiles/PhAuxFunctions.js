/ * P h A u x F u n c t i o n s . j s  
 * 	 C o n t a i n s   a l l   o f   t h e   r a n d o m   f u n c t i o n s   t h a t   t h e   m a i n   p r o g r a m  
 * 	 ( i n   P h P a r s e r . j s )   h a s   t o   c a l l   t o   a c c o m p l i s h   i t s   j o b .  
 *  
 * 	 G l o b a l   V a r i a b l e s   D e f i n e d :   N o n e  
 * 	 G l o b a l   V a r i a b l e s   U s e d :   N o n e  
 * /  
  
 / * F i n d I n v o i c e S e c t i o n ( i v d a t a )  
 * 	 i v d a t a :   a   r e f e r e n c e   t o   t h e   i n v o i c e   s h e e t  
 *  
 * 	 G l o b a l   V a r i a b l e s   U s e d :   i  
 *  
 * 	 F i n d   a   c o n t r a c t   h e a d e r   i n   t h e   i n v o i c e   a f t e r   t h e   c u r r e n t   c u r s o r   p o s i t i o n .  
 * 	 I f   i t   w r a p s   a r o u n d ,   w e ' v e   h i t   t h e   e n d   o f   t h e   d a t a ,   s o   j u s t   r e t u r n .  
 * 	 E x t r a c t   t h e   p h o n e   n u m b e r   f o r   t h i s   s e c t i o n ,   a n d   t h e n   l o o p   t o   f i n d   t h e  
 * 	 b e g i n n i n g   o f   t h e   c a l l   d a t a .   I f   t h e r e   i s   n o   c a l l   d a t a ,   f i n d   t h e   n e x t   h e a d e r  
 * 	 a n d   r e p e a t .  
 *  
 * 	 R e t u r n s   a   s t r i n g   c o n t a i n i n g   t h e   p h o n e   n u m b e r   f o r   t h e   n e x t   s e c t i o n ,   o r   a n   e m p t y  
 * 	 s t r i n g   i f   t h e   e n d   o f   t h e   d a t a   h a s   b e e n   r e a c h e d .   A l s o ,   s e t s   t h e   g l o b a l   c u r s o r   ' i '  
 * 	 a t   t h e   b e g i n n i n g   o f   t h e   c a l l   d a t a   f o r   t h a t   s e c t i o n .  
 * /  
  
 f u n c t i o n   F i n d I n v o i c e S e c t i o n ( i v d a t a ) {  
 	 v a r   t n u m   =   " " ;  
  
 	 d o {  
 	 	 v a r   h c e l l   =   i v d a t a . C o l u m n s ( 1 ) . F i n d ( " >=B@0:B  !" , i v d a t a . C e l l s ( i , 1 ) ) ;  
 	 	 v a r   c u r p o s   =   p a r s e I n t ( h c e l l . A d d r e s s . s p l i t ( " $ " ) [ 2 ] ) ;  
 	 	 i f ( c u r p o s < i )   r e t u r n   " " ;   / / e n d   o f   d a t a  
 	 	 i   =   c u r p o s + 5 0 ;   / / h e a d e r s   a r e   a t   l e a s t   t h a t   l o n g ;   j u s t   s a v e   s o m e   t i m e .  
  
 	 	 t n u m   =   ( / :   ( \ d + ) / ) . e x e c ( h c e l l . T e x t ) [ 1 ] ;   / /   n u m b e r   f o r   t h a t   S e c t i o n  
  
 	 	 h c e l l   =   i v d a t a . C o l u m n s ( 7 ) . F i n d ( " : " , i v d a t a . C e l l s ( i , 7 ) ) ;   / / l o o k   f o r   s t a r t   o f   c a l l   d a t a ;   w e   d o n ' t   j u s t   l o o p   o v e r   e v e r y   r o w   a n d   c h e c k   f o r  
 	 	 c u r p o s   =   p a r s e I n t ( h c e l l . A d d r e s s . s p l i t ( " $ " ) [ 2 ] ) ; 	 	 / / n e w   h e a d e r s   a t   t h e   s a m e   t i m e   b e c a u s e   t h a t   w o u l d   m i s s   t h e   c a s e   w h e r e   t h e   l a s t  
 	 	 i f ( c u r p o s < i )   r e t u r n   " " ;   / / e n d   o f   d a t a 	 	 	 / / n u m b e r   i n   t h e   i n v o i c e   h a s   n o   d a t a .   S o ,   j u s t   i n   c a s e ,   w e   c h e c k   t h a t   f i r s t .  
  
 	 	 w h i l e ( i v d a t a . C e l l s ( i , 1 ) . T e x t   ! =   " " )   i + + ;   / /   S k i p   t h e   r e m a i n i n g   h e a d e r   j u n k  
 	 	 w h i l e ( i < c u r p o s ) {   / / m a k e   s u r e   t h e r e   i s n ' t   a n o t h e r   h e a d e r   f i r s t  
 	 	 	 i f ( i v d a t a . C e l l s ( i , 1 ) . T e x t   ! =   " " ) {   / / i f   t h e r e   i s ,   t h e n   b r e a k  
 	 	 	 	 t n u m   =   " " ;  
 	 	 	 	 i - - ;  
 	 	 	 	 b r e a k ;  
 	 	 	 }  
 	 	 	 i + + ;  
 	 	 }  
 	 }   w h i l e ( t n u m   = =   " " ) ;   / /   i f   a   n e w   h e a d e r   s h o w e d   u p ,   w e   s e a r c h   a g a i n  
 	 r e t u r n   t n u m ;  
 }  
 